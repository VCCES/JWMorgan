Attribute VB_Name = "JankyBPPSolver"
'This work is licensed under the Creative Commons Attribution 4.0 International License. To view a copy of this license, visit http://creativecommons.org/licenses/by/4.0/.

Option Explicit

Global SteelMode As Boolean

Const epsilon As Double = 0.0001
Public Const offset_constant As Long = 21
'data declarations

Private Type item_type_data
    id As Long
    Width As Double
    Height As Double
    Area As Double
    rotatable As Boolean
    mandatory As Long
    profit As Double
    number_requested As Long
    sort_criterion As Double        ''' value equal to the sort criteria field's value
    Placement As String
    Size As String
End Type

Private Type item_list_data
    num_item_types As Long
    total_number_of_items As Long
    item_types() As item_type_data
End Type

Dim item_list As item_list_data

Private Type bin_type_data
    type_id As Long
    bin_name As String
    Width As Double
    Height As Double
    Area As Double
    mandatory As Long
    cost As Double
    number_available As Long
    Placement As String
    Size As String
End Type

Private Type bin_list_data
    num_bin_types As Long
    bin_types() As bin_type_data
End Type

Dim bin_list As bin_list_data

Private Type compatibility_data
    item_to_item() As Boolean
    bin_to_item() As Boolean
End Type

Dim compatibility_list As compatibility_data

Private Type item_location
    sw_x As Double
    sw_y As Double
    max_x As Double ' for Guillotine cuts
    max_y As Double ' for Guillotine cuts
End Type

Private Type item_in_bin
    item_type As Long
    item_name As String
    rotated As Boolean
    mandatory As Long
    sw_x As Double
    sw_y As Double
    ne_x As Double
    ne_y As Double
    max_x As Double ' for Guillotine cuts
    max_y As Double ' for Guillotine cuts
    first_cut_direction As Long ' for Guillotine cuts: 0 means horizontal, 1 means vertical
    cut_length As Double ' for Guillotine cuts
    Placement As String
    Size As String
End Type

Private Type bin_data
    type_id As Long
    bin_name As String
    Width As Double
    Height As Double
    Area As Double
    cost As Double
    item_cnt As Long
    mandatory As Long
    items() As item_in_bin
    addition_points() As item_location
    repack_item_count() As Long
    area_packed As Double
End Type

Private Type solution_data
    num_bins As Long
    feasible As Boolean
    net_profit As Double
    total_area As Double
    total_distance As Double
    total_area_utilization As Double
    item_type_order() As Long
    rotation_order() As Long
    first_cut_direction() As Long
    bin() As bin_data
    unpacked_item_count() As Long
End Type

Private Type instance_data
    item_item_compatibility_worksheet As Boolean 'true if the data exists
    bin_item_compatibility_worksheet As Boolean 'true if the data exists
    guillotine_cuts As Boolean
    global_upper_bound As Double        ''' max profit obtainable
End Type

Dim instance As instance_data

Private Type solver_option_data
    CPU_time_limit As Double
    item_sort_criterion As Long
    show_progress As Boolean
End Type

Dim solver_options As solver_option_data

Private Sub SortBins(solution As solution_data)
    
    Dim i As Long
    Dim j As Long
    Dim candidate_index As Long
    Dim max_mandatory As Long
    Dim max_area_packed As Double
    Dim min_ratio As Double
    Dim swap_bin As bin_data
        
    'insertion sort

    If Rnd < 0.8 Then
        
        'insertion sort
    
        With solution
    
            For i = 1 To .num_bins
                candidate_index = i
                max_mandatory = .bin(i).mandatory
                max_area_packed = .bin(i).area_packed
                min_ratio = .bin(i).cost / .bin(i).Area
    
                For j = i + 1 To .num_bins
    
                    If (.bin(j).mandatory > max_mandatory) Or _
                        ((.bin(j).mandatory = max_mandatory) And (.bin(j).area_packed > max_area_packed + epsilon)) Or _
                        ((.bin(j).mandatory = 0) And (max_mandatory = 0) And (.bin(j).area_packed > max_area_packed - epsilon) And ((.bin(j).cost / .bin(j).Area) < min_ratio)) Then
    
                        candidate_index = j
                        max_mandatory = .bin(j).mandatory
                        max_area_packed = .bin(j).area_packed
                        min_ratio = .bin(j).cost / .bin(j).Area
    
                    End If
    
                Next j
    
                If candidate_index <> i Then
                    swap_bin = .bin(candidate_index)
                    .bin(candidate_index) = .bin(i)
                    .bin(i) = swap_bin
                End If
    
            Next i

        End With
    Else
        
        With solution
            For i = 1 To .num_bins
                
                candidate_index = Int((.num_bins - i + 1) * Rnd + i)
    
                If candidate_index <> i Then
                    swap_bin = .bin(candidate_index)
                    .bin(candidate_index) = .bin(i)
                    .bin(i) = swap_bin
                End If
    
            Next i
        End With
        
    End If
    
End Sub
Private Sub PerturbSolution(solution As solution_data)
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Dim swap_long As Long
    
    Dim bin_emptying_probability As Double
    Dim item_removal_probability As Double
    Dim repack_flag As Boolean
    Dim continue_flag As Boolean
    
    Dim empty_type_probability As Double
    Dim empty_type As Long
    
    empty_type_probability = Rnd
    If empty_type_probability < 0.5 Then
        empty_type = 0
    Else
        empty_type = 1
    End If
    
    For i = 1 To solution.num_bins
    
        With solution.bin(i)
        
            If empty_type = 0 Then
                bin_emptying_probability = 1 - 0.8 * (.area_packed / .Area)
                item_removal_probability = 0
            ElseIf empty_type = 1 Then
                bin_emptying_probability = 0.2
                item_removal_probability = 0.1 + Rnd * 0.1
            End If
            
            If .item_cnt > 0 Then
            
                If Rnd() < bin_emptying_probability Then
                    
                    'empty the bin
                    
                    For j = 1 To .item_cnt
                    
                        solution.unpacked_item_count(.items(j).item_type) = solution.unpacked_item_count(.items(j).item_type) + 1
                        solution.net_profit = solution.net_profit - item_list.item_types(.items(j).item_type).profit
                        
                    Next j
                    solution.net_profit = solution.net_profit + .cost
                    solution.total_area = solution.total_area - .area_packed
                    
                    .item_cnt = 0
                    .area_packed = 0
                    .addition_points(1).sw_x = 0
                    .addition_points(1).sw_y = 0
                    .addition_points(1).max_x = .Width
                    .addition_points(1).max_y = .Height
                    
                Else
                
                    repack_flag = False

                    For j = 1 To .item_cnt

                        If ((solution.feasible = False) And (.items(j).mandatory = 0)) Or (Rnd() < item_removal_probability) Then

                            solution.unpacked_item_count(.items(j).item_type) = solution.unpacked_item_count(.items(j).item_type) + 1

                            solution.net_profit = solution.net_profit - item_list.item_types(.items(j).item_type).profit

                            .items(j).item_type = 0

                            repack_flag = True
                        End If

                    Next j

                    If repack_flag = True Then

                        For j = 1 To .item_cnt

                            If .items(j).item_type > 0 Then
                                solution.net_profit = solution.net_profit - item_list.item_types(.items(j).item_type).profit
                            End If

                        Next j
                        solution.net_profit = solution.net_profit + .cost
                        solution.total_area = solution.total_area - .area_packed

                        For j = 1 To item_list.num_item_types
                            .repack_item_count(j) = 0
                        Next j

                        For j = 1 To .item_cnt

                            If .items(j).item_type > 0 Then
                                .repack_item_count(.items(j).item_type) = .repack_item_count(.items(j).item_type) + 1
                            End If

                        Next j

                        .area_packed = 0
                        .item_cnt = 0
                        .addition_points(1).sw_x = 0
                        .addition_points(1).sw_y = 0
                        .addition_points(1).max_x = .Width
                        .addition_points(1).max_y = .Height

                        'repack now

                        For j = 1 To item_list.num_item_types

                            continue_flag = True
                            Do While (.repack_item_count(j) > 0) And (continue_flag = True)
                                continue_flag = AddItemToBin(solution, i, j, 2)
                            Loop

                            ' put the remaining items in the unpacked items list

                            solution.unpacked_item_count(j) = solution.unpacked_item_count(j) + .repack_item_count(j)
                            .repack_item_count(j) = 0

                        Next j

                    End If
                
                End If
            
            End If
            
        End With
        
    Next i
    
    'change the preferred rotation order randomly
    
    For i = 1 To item_list.num_item_types

        If Rnd() < 0.2 Then
            If solution.rotation_order(i, 1) = 1 Then
                solution.rotation_order(i, 1) = 0
                solution.rotation_order(i, 2) = 1
            Else
                solution.rotation_order(i, 1) = 1
                solution.rotation_order(i, 2) = 0
            End If
        End If

    Next i
    
    'change the first cut direction randomly
    
    For i = 1 To item_list.num_item_types

        If Rnd() < 0.2 Then
            If solution.first_cut_direction(i) = 1 Then
                solution.first_cut_direction(i) = 0
            Else
                solution.first_cut_direction(i) = 1
            End If
        End If

    Next i
    
    'change the item order randomly

    For i = 1 To item_list.num_item_types

        If Rnd < 0.1 Then
            j = Int((item_list.num_item_types - i + 1) * Rnd + i) ' the order to swap with
    
            swap_long = solution.item_type_order(i)
            solution.item_type_order(i) = solution.item_type_order(j)
            solution.item_type_order(j) = swap_long
        End If
    Next i
    
End Sub
Private Function AddItemToBin(solution As solution_data, bin_index As Long, item_type_index As Long, add_type As Long)
        
    Dim i As Long
    Dim j As Long
    Dim rotation As Long
    
    Dim sw_x As Double
    Dim sw_y As Double
    Dim ne_x As Double
    Dim ne_y As Double
    
    Dim min_x As Double
    Dim min_y As Double
    Dim candidate_position As Double
    Dim candidate_rotation As Long
    
    With solution.bin(bin_index)
    
        min_x = .Width + 1
        min_y = .Height + 1
        candidate_position = 0
        
        
        
        'area size check
        If .area_packed + item_list.item_types(item_type_index).Area > .Area Then GoTo AddItemToBin_Finish
        
        'item to item compatibility check
        For rotation = 1 To 2
        
            If (solution.rotation_order(item_type_index, rotation) = 1) And (item_list.item_types(item_type_index).rotatable = False) Then
                GoTo NextRotation
            End If

            For i = 1 To .item_cnt + 1
                
                sw_x = .addition_points(i).sw_x
                sw_y = .addition_points(i).sw_y
                
                If solution.rotation_order(item_type_index, rotation) = 0 Then
                    ne_x = sw_x + item_list.item_types(item_type_index).Width
                    ne_y = sw_y + item_list.item_types(item_type_index).Height
                Else
                    ne_x = sw_x + item_list.item_types(item_type_index).Height
                    ne_y = sw_y + item_list.item_types(item_type_index).Width
                End If
                
                'check the feasibility of all four corners, w.r.t to the other items
                
                If (ne_x > .Width + epsilon) Or (ne_y > .Height + epsilon) Then GoTo NextIteration
                
                If instance.guillotine_cuts = True Then
                    If (ne_x > .addition_points(i).max_x + epsilon) Or (ne_y > .addition_points(i).max_y + epsilon) Then GoTo NextIteration
                End If
                
                For j = 1 To .item_cnt
                    
                    If (sw_x < .items(j).ne_x - epsilon) And (ne_x > .items(j).sw_x + epsilon) And (ne_y > .items(j).sw_y + epsilon) And (sw_y < .items(j).ne_y - epsilon) Then GoTo NextIteration
                
                Next j
                
                
                'no conflicts at this point
                
                If (sw_y < min_y) Or _
                  ((sw_y <= min_y + epsilon) And (sw_x < min_x)) Then
                   min_x = sw_x
                   min_y = sw_y
                   candidate_position = i
                   candidate_rotation = solution.rotation_order(item_type_index, rotation)
                End If
NextIteration:
            Next i
            
NextRotation:
        Next rotation
        
    End With
    
AddItemToBin_Finish:

    If candidate_position = 0 Then
        AddItemToBin = False
    Else
        With solution.bin(bin_index)
            .item_cnt = .item_cnt + 1
            .items(.item_cnt).item_type = item_type_index
            .items(.item_cnt).sw_x = .addition_points(candidate_position).sw_x
            .items(.item_cnt).sw_y = .addition_points(candidate_position).sw_y
            If candidate_rotation = 1 Then
                .items(.item_cnt).rotated = True
            Else
                .items(.item_cnt).rotated = False
            End If
            .items(.item_cnt).mandatory = item_list.item_types(item_type_index).mandatory
            .items(.item_cnt).Placement = item_list.item_types(item_type_index).Placement
            .items(.item_cnt).Size = item_list.item_types(item_type_index).Size
            If candidate_rotation = 0 Then
                .items(.item_cnt).ne_x = .items(.item_cnt).sw_x + item_list.item_types(item_type_index).Width
                .items(.item_cnt).ne_y = .items(.item_cnt).sw_y + item_list.item_types(item_type_index).Height
            Else
                .items(.item_cnt).ne_x = .items(.item_cnt).sw_x + item_list.item_types(item_type_index).Height
                .items(.item_cnt).ne_y = .items(.item_cnt).sw_y + item_list.item_types(item_type_index).Width
            End If
            
            .area_packed = .area_packed + item_list.item_types(item_type_index).Area

            If instance.guillotine_cuts = True Then
                .items(.item_cnt).first_cut_direction = solution.first_cut_direction(item_type_index)
                
                If ((.items(.item_cnt).first_cut_direction = 0) And (.items(.item_cnt).ne_y = .addition_points(candidate_position).max_y)) Or ((.items(.item_cnt).first_cut_direction = 1) And (.items(.item_cnt).ne_x = .addition_points(candidate_position).max_x)) Then
                    .items(.item_cnt).cut_length = 0
                Else
                    .items(.item_cnt).cut_length = 0
                    If .items(.item_cnt).first_cut_direction = 0 Then
                        If .items(.item_cnt).ne_y < .addition_points(candidate_position).max_y - epsilon Then
                            .items(.item_cnt).cut_length = .items(.item_cnt).cut_length + (.addition_points(candidate_position).max_x - .items(.item_cnt).sw_x)
                            '.items(.item_cnt).cut_length = .items(.item_cnt).cut_length + 1
                        End If

                        If .items(.item_cnt).ne_x < .addition_points(candidate_position).max_x - epsilon Then
                            .items(.item_cnt).cut_length = .items(.item_cnt).cut_length + (.items(.item_cnt).ne_y - .items(.item_cnt).sw_y)
                            '.items(.item_cnt).cut_length = .items(.item_cnt).cut_length + 1
                        End If
                    Else
                        If .items(.item_cnt).ne_x < .addition_points(candidate_position).max_x - epsilon Then
                            .items(.item_cnt).cut_length = .items(.item_cnt).cut_length + (.addition_points(candidate_position).max_y - .items(.item_cnt).sw_y)
                            '.items(.item_cnt).cut_length = .items(.item_cnt).cut_length + 1
                        End If
                        If .items(.item_cnt).ne_y < .addition_points(candidate_position).max_y - epsilon Then
                            .items(.item_cnt).cut_length = .items(.item_cnt).cut_length + (.items(.item_cnt).ne_x - .items(.item_cnt).sw_x)
                            '.items(.item_cnt).cut_length = .items(.item_cnt).cut_length + 1
                        End If
                    End If
                 End If
                
                .items(.item_cnt).max_x = .addition_points(candidate_position).max_x
                .items(.item_cnt).max_y = .addition_points(candidate_position).max_y
            End If
            

            If add_type = 2 Then
                .repack_item_count(item_type_index) = .repack_item_count(item_type_index) - 1
            End If
            
            'update the addition points
            
            For i = candidate_position To .item_cnt - 1
                .addition_points(i) = .addition_points(i + 1)
            Next i
            
            .addition_points(.item_cnt).sw_x = .items(.item_cnt).ne_x
            .addition_points(.item_cnt).sw_y = .items(.item_cnt).sw_y
            
            .addition_points(.item_cnt + 1).sw_x = .items(.item_cnt).sw_x
            .addition_points(.item_cnt + 1).sw_y = .items(.item_cnt).ne_y
            
            If instance.guillotine_cuts = True Then
                If .items(.item_cnt).first_cut_direction = 0 Then
                    .addition_points(.item_cnt).max_x = .items(.item_cnt).max_x
                    .addition_points(.item_cnt).max_y = .items(.item_cnt).ne_y
                    
                    .addition_points(.item_cnt + 1).max_x = .items(.item_cnt).max_x
                    .addition_points(.item_cnt + 1).max_y = .items(.item_cnt).max_y
                Else
                    .addition_points(.item_cnt).max_x = .items(.item_cnt).max_x
                    .addition_points(.item_cnt).max_y = .items(.item_cnt).max_y
                    
                    .addition_points(.item_cnt + 1).max_x = .items(.item_cnt).ne_x
                    .addition_points(.item_cnt + 1).max_y = .items(.item_cnt).max_y
                End If
            End If
            
        End With
        
        With solution
            'update the profit
            
            If .bin(bin_index).item_cnt = 1 Then
                .net_profit = .net_profit + item_list.item_types(item_type_index).profit - .bin(bin_index).cost
            Else
                .net_profit = .net_profit + item_list.item_types(item_type_index).profit
            End If
            
            'update the area per bin and the total area
            
            .total_area = .total_area + item_list.item_types(item_type_index).Area
            
            'update the unpacked items
            
            If add_type = 1 Then
                .unpacked_item_count(item_type_index) = .unpacked_item_count(item_type_index) - 1
            End If
            
        End With
        
        AddItemToBin = True
    End If
    
End Function



'''''''''''''''''''''' Sub for reading data from the item sheet, writing to item type
Private Sub GetItemData(InputCollection As Collection, Optional SteelClass As Boolean)
    
    item_list.num_item_types = InputCollection.Count
    item_list.total_number_of_items = 0
    
    ReDim item_list.item_types(1 To item_list.num_item_types)
    
    Dim i As Long
    
    With item_list
        
        For i = 1 To .num_item_types
            
            .item_types(i).id = i
            .item_types(i).Width = 1
            If SteelClass = False Then
                .item_types(i).Height = InputCollection(i).tLength
                .item_types(i).Area = InputCollection(i).tLength
                .item_types(i).number_requested = InputCollection(i).Quantity
            Else
                .item_types(i).Height = InputCollection(i).Length
                .item_types(i).Area = InputCollection(i).Length
                .item_types(i).number_requested = InputCollection(i).Qty
                .item_types(i).Placement = InputCollection(i).Placement
                .item_types(i).Size = InputCollection(i).Size
            End If

            
            

            .item_types(i).rotatable = False

            
            'make all items mandatory
            .item_types(i).mandatory = 1
            
            .item_types(i).profit = .item_types(i).Height
            
            
            
            If solver_options.item_sort_criterion = 1 Then
                .item_types(i).sort_criterion = .item_types(i).Area
            ElseIf solver_options.item_sort_criterion = 2 Then
                .item_types(i).sort_criterion = .item_types(i).Width + .item_types(i).Height
            ElseIf solver_options.item_sort_criterion = 3 Then
                .item_types(i).sort_criterion = .item_types(i).Height
            ElseIf solver_options.item_sort_criterion = 4 Then
                .item_types(i).sort_criterion = .item_types(i).Width
            End If
            
            item_list.total_number_of_items = item_list.total_number_of_items + .item_types(i).number_requested
        
        Next i
    
    End With
    
End Sub


Private Sub InitializeSolution(solution As solution_data)
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    
    With solution
        .feasible = False
        .net_profit = 0
        .total_area = 0
        .total_distance = 0
        .total_area_utilization = 0
        
        .num_bins = 0
        For i = 1 To bin_list.num_bin_types
            If bin_list.bin_types(i).mandatory >= 0 Then
                .num_bins = .num_bins + bin_list.bin_types(i).number_available
            End If
        Next i
        
        ReDim .item_type_order(1 To item_list.num_item_types)
        For i = 1 To item_list.num_item_types
            .item_type_order(i) = i
        Next i
        
        ReDim .rotation_order(1 To item_list.num_item_types, 1 To 2)
        For i = 1 To item_list.num_item_types
            .rotation_order(i, 1) = 0
            .rotation_order(i, 2) = 1
        Next i
        
        ReDim .first_cut_direction(1 To item_list.num_item_types)
        For i = 1 To item_list.num_item_types
            .first_cut_direction(i) = 0
        Next i
        
        ReDim .bin(1 To .num_bins)
        For i = 1 To .num_bins
            ReDim .bin(i).items(1 To 2 * item_list.total_number_of_items)
            ReDim .bin(i).addition_points(1 To item_list.total_number_of_items + 1)
            ReDim .bin(i).repack_item_count(1 To item_list.total_number_of_items)
        Next i
        
        ReDim .unpacked_item_count(1 To item_list.num_item_types)
        
        l = 1
        For i = 1 To bin_list.num_bin_types
            If bin_list.bin_types(i).mandatory >= 0 Then
                For j = 1 To bin_list.bin_types(i).number_available
                    
                    .bin(l).Width = bin_list.bin_types(i).Width
                    .bin(l).Height = bin_list.bin_types(i).Height
                    .bin(l).Area = bin_list.bin_types(i).Area
                    .bin(l).cost = bin_list.bin_types(i).cost
                    .bin(l).mandatory = bin_list.bin_types(i).mandatory
                    .bin(l).type_id = i
                    .bin(l).area_packed = 0
                    .bin(l).item_cnt = 0
                    
                    For k = 1 To item_list.total_number_of_items
                        .bin(l).items(k).item_type = 0
                        .bin(l).addition_points(k).sw_x = 0
                        .bin(l).addition_points(k).sw_y = 0
                        .bin(l).addition_points(k).max_x = .bin(l).Width
                        .bin(l).addition_points(k).max_y = .bin(l).Height
                    Next k
                    
                    For k = 1 To item_list.total_number_of_items
                        .bin(l).repack_item_count(k) = 0
                    Next k
                    
                    l = l + 1
                Next j
            End If
        Next i
        
        For i = 1 To item_list.num_item_types
            .unpacked_item_count(i) = item_list.item_types(i).number_requested
        Next i
        
    End With
    
End Sub


Private Sub WriteSolution(solution As solution_data, SolvedCollection As Collection, InputType As String)
Dim TrimPiece As clsTrim
Dim CurrentBin As Integer
Dim Member As clsMember
Dim Span As clsMember
   
    Application.Calculation = xlCalculationManual
            
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim m As Long
    Dim BinIndex As Integer
    
    Dim bin_index As Long
    
    Dim swap_bin As bin_data
    
    'my vars
    Dim BinName As String
    Dim ItemName As String
    Dim BinNumber As Integer
    
    'reset bin index
    BinIndex = 1
    
    'sort the bins
    'first ensure that simlar patterns occur together
    
    For i = 1 To solution.num_bins
        solution.bin(i).area_packed = solution.bin(i).area_packed * solution.bin(i).Area * solution.bin(i).Area
        For j = 1 To solution.bin(i).item_cnt
            solution.bin(i).area_packed = solution.bin(i).area_packed + item_list.item_types(solution.bin(i).items(j).item_type).Area * item_list.item_types(solution.bin(i).items(j).item_type).Area
            Debug.Print item_list.item_types(solution.bin(i).items(j).item_type).Area
            Debug.Print item_list.item_types(solution.bin(i).items(j).item_type).id
            Debug.Print item_list.item_types(solution.bin(i).items(j).item_type).number_requested
        Next j
    Next i
    
    For i = 1 To solution.num_bins
        For j = solution.num_bins To 2 Step -1
            If (solution.bin(j).type_id < solution.bin(j - 1).type_id) Or _
                ((solution.bin(j).type_id = solution.bin(j - 1).type_id) And (solution.bin(j).area_packed > solution.bin(j - 1).area_packed)) Then
                swap_bin = solution.bin(j)
                solution.bin(j) = solution.bin(j - 1)
                solution.bin(j - 1) = swap_bin
            End If
        Next j
    Next i
    
    If solution.feasible = False And InputType <> "Roof Purlin" Then
        MsgBox "Warning: Last solution returned by the solver does not satisfy all constraints. The parts were not optimized, partially optimized, or no optimization was possible. Please check " & InputType & " parts."
    End If
    
    bin_index = 1
    
    'Expanded solution output - test
    With solution
        bin_index = 1
        CurrentBin = 0
        For i = 1 To .num_bins
            If .bin(i).item_cnt > 0 Then
                Debug.Print "Bin #: " & i & " " & .bin(i).Height
                Debug.Print "item count: " & .bin(i).item_cnt
                For j = 1 To .bin(i).item_cnt
                    If j > 1 Then
                        Debug.Print .bin(i).items(j).ne_y - .bin(i).items(j).sw_y
                    Else
                        Debug.Print .bin(i).items(j).ne_y
                    End If
                Next j
            End If
        Next i
    End With
    
    
    With solution
    
        'condensed solution output
        bin_index = 1
        CurrentBin = 0
        For i = 1 To bin_list.num_bin_types
            For j = 1 To bin_list.bin_types(i).number_available
                'name bins
                Select Case .bin(bin_index).Height
                ''' steel
                Case 20 * 12
                    .bin(bin_index).bin_name = "20'"
                Case 25 * 12
                    .bin(bin_index).bin_name = "25'"
                Case 30 * 12
                    .bin(bin_index).bin_name = "30'"
                '''''''''''' trim ''''''''
                Case 42
                    .bin(bin_index).bin_name = "3'6"""
                Case 75
                    .bin(bin_index).bin_name = "6'3"""
                Case 86
                    .bin(bin_index).bin_name = "7'2"""
                Case 87
                    .bin(bin_index).bin_name = "7'3"""
                Case 99
                    .bin(bin_index).bin_name = "8'3"""
                Case 122
                    .bin(bin_index).bin_name = "10'2"""
                Case 123
                    .bin(bin_index).bin_name = "10'3"""
                Case 146
                    .bin(bin_index).bin_name = "12'2"""
                Case 147
                    .bin(bin_index).bin_name = "12'3"""
                Case 170
                    .bin(bin_index).bin_name = "14'2"""
                Case 171
                    .bin(bin_index).bin_name = "14'3"""
                Case 194
                    .bin(bin_index).bin_name = "16'2"""
                Case 195
                    .bin(bin_index).bin_name = "16'3"""
                Case 218
                    .bin(bin_index).bin_name = "18'2"""
                Case 219
                    .bin(bin_index).bin_name = "18'3"""
                Case 244
                    .bin(bin_index).bin_name = "20'4"""
                End Select
                If bin_list.bin_types(i).mandatory >= 0 Then
                    'if items in the bin, add to the output bin number
                    If .bin(bin_index).item_cnt <> 0 Then
                        BinNumber = BinNumber + 1
                        'add previous trim piece
                        'If BinNumber <> 1 Then SolvedCollection.Add TrimPiece
                    End If
                    'check for items
                    For k = 1 To .bin(bin_index).item_cnt
                        If CurrentBin <> BinNumber Then
                            If SteelMode = False Then
                                'new trim class
                                Set TrimPiece = New clsTrim
                                TrimPiece.tMeasurement = .bin(bin_index).bin_name
                                TrimPiece.tLength = .bin(bin_index).Height
                                'girt whatever.placement = .bin(bin_index).placement
                                Select Case InputType
                                Case "Jamb"
                                    TrimPiece.tType = "Jamb Trim"
                                Case "Head"
                                    TrimPiece.tType = "Head Trim W/ Kickout"
                                End Select
                                TrimPiece.Quantity = 1
                                SolvedCollection.Add TrimPiece
                            Else
                                'new member class
                                Set Member = New clsMember
                                Member.Measurement = .bin(bin_index).bin_name
                                Member.Length = .bin(bin_index).Height ' This is the total length (ie - 20' or 25' or 30')
                                Member.Size = .bin(bin_index).items(1).Size
                                If InputType = "Roof Purlin" Then
                                    Member.Placement = InputType & " " & Member.Size & " Span # " & BinIndex
                                Else
                                    Member.Placement = Member.Size & " Span # " & BinIndex
                                End If
                                BinIndex = BinIndex + 1
                                'girt whatever.placement = .bin(bin_index).placement
                                Select Case InputType
                                Case "Girt"
                                    Member.mType = "C Purlin"
                                    'Member.Size = "8"" C Purlin"
                                Case "Roof Purlin"
                                    Member.mType = "C Purlin"
                                Case "TS"
                                    Member.mType = "Tube Steel"
                                Case "IBeam"
                                    Member.mType = "I-Beam"
                                End Select
                                Member.Qty = 1
                                'debug output
                                If SteelMode = True Then
                                For m = 1 To .bin(bin_index).item_cnt
                                    If m > 1 Then
                                        Set Span = New clsMember
                                        Span.Length = .bin(bin_index).items(m).ne_y - .bin(bin_index).items(m).sw_y
                                        Span.Qty = 1
                                        Span.Placement = .bin(bin_index).items(m).Placement
                                        Span.Size = .bin(bin_index).items(m).Size
                                        Member.ComponentMembers.Add Span
                                        'Debug.Print "Added Member No.: " & k & ", Member length: "; .bin(bin_index).items(k).ne_y - .bin(bin_index).items(k - 1).ne_y & ", Placement: " & .bin(bin_index).items(k).Placement
                                    Else
                                        Set Span = New clsMember
                                        Span.Length = .bin(bin_index).items(m).ne_y
                                        Span.Qty = 1
                                        Span.Placement = .bin(bin_index).items(m).Placement
                                        Span.Size = .bin(bin_index).items(m).Size
                                        Member.ComponentMembers.Add Span
                                        'Debug.Print "Bin # " & bin_index & ", Order Length: " & .bin(bin_index).bin_name
                                        'Debug.Print "Added Member No.: " & k & ", Member length: "; .bin(bin_index).items(k).ne_y & ", Placement: " & .bin(bin_index).items(k).Placement
                                    End If
                                Next m
                                End If
                                SolvedCollection.Add Member
                            End If
                            
                                                            
                            'update current bin
                            CurrentBin = BinNumber
                        End If
                        
                        
                        
                    Next k
                    bin_index = bin_index + 1
                End If
            Next j
        Next i
        
    End With
    

        
    
    Application.Calculation = xlCalculationAutomatic
    
End Sub

Private Sub ReadSolution(solution As solution_data)
       
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
            
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    
    Dim bin_index As Long
    Dim item_type_index As Long
    
    Dim offset As Long
    
    offset = 0
    bin_index = 1
    
    With solution
    
        For i = 1 To bin_list.num_bin_types
        
            For j = 1 To bin_list.bin_types(i).number_available
                    
                If bin_list.bin_types(i).mandatory >= 0 Then
                    
                    With .bin(bin_index)
                        
                        l = Cells(4, offset + 17).Value
                        
                        For k = 1 To l
                            If IsNumeric(Cells(5 + k, offset + 14).Value) = True Then
                            
                                .item_cnt = .item_cnt + 1
                                
                                item_type_index = Cells(5 + k, offset + 14).Value
                                
                                solution.unpacked_item_count(item_type_index) = solution.unpacked_item_count(item_type_index) - 1
                                
                                .items(.item_cnt).item_type = item_type_index
                                .items(.item_cnt).sw_x = Cells(5 + k, offset + 3).Value
                                .items(.item_cnt).sw_y = Cells(5 + k, offset + 4).Value
                                  
                                If ThisWorkbook.Worksheets("3.Solution").Cells(5 + k, offset + 5).Value = "Yes" Then
                                    .items(.item_cnt).rotated = True
                                    .items(.item_cnt).ne_x = .items(.item_cnt).sw_x + item_list.item_types(item_type_index).Height
                                    .items(.item_cnt).ne_y = .items(.item_cnt).sw_y + item_list.item_types(item_type_index).Width
                                Else
                                    .items(.item_cnt).rotated = False
                                    .items(.item_cnt).ne_x = .items(.item_cnt).sw_x + item_list.item_types(item_type_index).Width
                                    .items(.item_cnt).ne_y = .items(.item_cnt).sw_y + item_list.item_types(item_type_index).Height
                                End If
                                
                                If instance.guillotine_cuts = True Then
                                    If ThisWorkbook.Worksheets("3.Solution").Cells(5 + k, offset + 6).Value = .items(.item_cnt).sw_x Then
                                        .items(.item_cnt).first_cut_direction = 0
                                        .items(.item_cnt).max_x = Cells(5 + k, offset + 8).Value
                                    Else
                                        .items(.item_cnt).first_cut_direction = 1
                                        .items(.item_cnt).max_y = Cells(5 + k, offset + 9).Value
                                    End If
                                    
                                End If
                        
                                .area_packed = .area_packed + item_list.item_types(item_type_index).Area
                                
                                If .item_cnt = 1 Then
                                    solution.net_profit = solution.net_profit + item_list.item_types(item_type_index).profit - .cost
                                Else
                                    solution.net_profit = solution.net_profit + item_list.item_types(item_type_index).profit
                                End If
                                
                            End If
                        Next k
                        
                    End With
                    
                    bin_index = bin_index + 1
                End If
                
                offset = offset + offset_constant
            Next j
        Next i
    End With
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub
Sub BPP_Solver(SolvedCollection As Collection, InputCollection As Collection, InputType As String, Optional FOType As String, Optional Wall As String)
Dim i As Long
Dim AvailableBins As Integer
Dim result As Long
Dim ItemCount As Integer
Dim TrimPiece As clsTrim
Dim TotalTrimLength As Integer
Dim TotalMemberLength As Double
Dim Member As clsMember


    'determine total item count
    If InputType <> "Girt" And InputType <> "Roof Purlin" And InputType <> "TS" And InputType <> "IBeam" Then
        For Each TrimPiece In InputCollection
            ItemCount = ItemCount + TrimPiece.Quantity
            TotalTrimLength = TotalTrimLength + (TrimPiece.tLength * TrimPiece.Quantity)
        Next TrimPiece
        Debug.Print FOType & " " & InputType & " " & TotalTrimLength
        SteelMode = False
    Else
        SteelMode = True
        For Each Member In InputCollection
            ItemCount = ItemCount + Member.Qty
            TotalMemberLength = TotalMemberLength + (Member.Length * Member.Qty)
        Next Member
        Debug.Print Wall & " " & InputType & " " & TotalMemberLength
    End If
        

'Instance data variables

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim WorksheetExists As Boolean
    Dim reply As Integer

    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo BPP_Solver_Finish
    
    '''''''''''''''''''''''''''//// Set Initial Solver Options ///
    With solver_options
        ''' Sort Criteria: 1 = Area, 2 = Circumference, 3 = height, 4 = Width
        ''' Show_Progress property = Display progress in status bar
        ''' CPU_Time Limit: Run Time in Seconds. Minimum: 1 second/item
        .item_sort_criterion = 3
        .show_progress = True
        'longer runtime for sections with bigger openings
        If FOType = "OHDoors" Or FOType = "MiscFOs" Or FOType = "" Then
            .CPU_time_limit = ItemCount * 5
        ElseIf InputType = "Girt" Then
            .CPU_time_limit = ItemCount / 5
        Else
            .CPU_time_limit = ItemCount / 5
        End If
    End With
          
    '''''''''''''''''''''''''''//// Set up "Items" (The Trim Lengths to be Ordered) ///
    If InputType = "Girt" Or InputType = "Roof Purlin" Or InputType = "TS" Or InputType = "IBeam" Then
        Call GetItemData(InputCollection, True)
    Else
        Call GetItemData(InputCollection)
    End If
    
    '''''''''''''''''''''''''''//// Set up "Bins" (Stock Standard Trim Sizes" ///
    ' Calculate Max Possible Bins Needed
    AvailableBins = 5     'Placeholder
    With bin_list
        'number of stock trim types
        If InputType = "Jamb" Then
            .num_bin_types = 7
        ElseIf InputType = "Head" Then
            .num_bin_types = 11
        ElseIf InputType = "Girt" Then
            .num_bin_types = 3
        ElseIf InputType = "Roof Purlin" Then
            .num_bin_types = 3
        ElseIf InputType = "TS" Then
            .num_bin_types = 3
        ElseIf InputType = "IBeam" Then
            .num_bin_types = 9
        End If
        ReDim .bin_types(1 To .num_bin_types)
        For i = 1 To .num_bin_types
            .bin_types(i).type_id = i
            .bin_types(i).Width = 1
            'Mandatory Values: 1 = Must use; 0 = May use; -1 = Don't use
            .bin_types(i).mandatory = 0
            'Number available: needs to be calculated to optimize (I think)

            If InputType = "Jamb" Then
                Select Case i
                Case 1
                    .bin_types(i).Height = 86
                    .bin_types(i).bin_name = "7'2"""
                Case 2
                    .bin_types(i).Height = 122
                    .bin_types(i).bin_name = "10'2"""
                Case 3
                    .bin_types(i).Height = 146
                    .bin_types(i).bin_name = "12'2"""
                Case 4
                    .bin_types(i).Height = 170
                    .bin_types(i).bin_name = "14'2"""
                Case 5
                    .bin_types(i).Height = 194
                    .bin_types(i).bin_name = "16'2"""
                Case 6
                    .bin_types(i).Height = 218
                    .bin_types(i).bin_name = "18'2"""
                Case 7
                    .bin_types(i).Height = 244
                    .bin_types(i).bin_name = "20'4"""
                End Select
            ElseIf InputType = "Head" Then
                Select Case i
                Case 1
                    .bin_types(i).Height = 42
                    .bin_types(i).bin_name = "3'6"""
                Case 2
                    .bin_types(i).Height = 75
                    .bin_types(i).bin_name = "6'3"""
                Case 3
                    .bin_types(i).Height = 87
                    .bin_types(i).bin_name = "7'3"""
                Case 4
                    .bin_types(i).Height = 99
                    .bin_types(i).bin_name = "8'3"""
                Case 5
                    .bin_types(i).Height = 122
                    .bin_types(i).bin_name = "10'2"""
                Case 6
                    .bin_types(i).Height = 123
                    .bin_types(i).bin_name = "10'3"""
                Case 7
                    .bin_types(i).Height = 147
                    .bin_types(i).bin_name = "12'3"""
                Case 8
                    .bin_types(i).Height = 171
                    .bin_types(i).bin_name = "14'3"""
                Case 9
                    .bin_types(i).Height = 195
                    .bin_types(i).bin_name = "16'3"""
                Case 10
                    .bin_types(i).Height = 219
                    .bin_types(i).bin_name = "18'3"""
                Case 11
                    .bin_types(i).Height = 244
                    .bin_types(i).bin_name = "20'4"""
                End Select
            ElseIf InputType = "Girt" Or InputType = "Roof Purlin" Then
                Select Case i
                Case 1
                    .bin_types(i).Height = 20 * 12
                    .bin_types(i).bin_name = "20'"
                Case 2
                    .bin_types(i).Height = 25 * 12
                    .bin_types(i).bin_name = "25'"
                Case 3
                    .bin_types(i).Height = 30 * 12
                    .bin_types(i).bin_name = "30'"
                End Select
            ElseIf InputType = "TS" Then
                Select Case i
                Case 1
                    
                    .bin_types(i).Height = 20 * 12
                    .bin_types(i).bin_name = "20'"
                Case 2
                    .bin_types(i).Height = 24 * 12
                    .bin_types(i).bin_name = "24'"
                Case 3
                    .bin_types(i).Height = 40 * 12
                    .bin_types(i).bin_name = "40'"
                
                
                
                End Select
            ElseIf InputType = "IBeam" Then
                Select Case i
                Case 1
                    .bin_types(i).Height = 20 * 12
                    .bin_types(i).bin_name = "20'"
                Case 2
                    .bin_types(i).Height = 25 * 12
                    .bin_types(i).bin_name = "25'"
                Case 3
                    .bin_types(i).Height = 30 * 12
                    .bin_types(i).bin_name = "30'"
                Case 4
                    .bin_types(i).Height = 35 * 12
                    .bin_types(i).bin_name = "35'"
                Case 5
                    .bin_types(i).Height = 40 * 12
                    .bin_types(i).bin_name = "40'"
                Case 6
                    .bin_types(i).Height = 45 * 12
                    .bin_types(i).bin_name = "45'"
                Case 7
                    .bin_types(i).Height = 50 * 12
                    .bin_types(i).bin_name = "50'"
                Case 8
                    .bin_types(i).Height = 55 * 12
                    .bin_types(i).bin_name = "55'"
                Case 9
                    .bin_types(i).Height = 60 * 12
                    .bin_types(i).bin_name = "60'"
                End Select
            End If
            'trick it into cost/area
            .bin_types(i).Area = .bin_types(i).Height
            .bin_types(i).cost = .bin_types(i).Height
            If InputType <> "Girt" And InputType <> "Roof Purlin" And InputType <> "TS" And InputType <> "IBeam" Then
                .bin_types(i).number_available = Application.WorksheetFunction.RoundUp(TotalTrimLength / .bin_types(i).Height, 0)
            Else
                .bin_types(i).number_available = Application.WorksheetFunction.RoundUp(TotalMemberLength / .bin_types(i).Height, 0)
            End If

            'lower number available if over item count
            If .bin_types(i).number_available > ItemCount Then .bin_types(i).number_available = ItemCount
            Debug.Print FOType & Wall & ", " & InputType & " - " & .bin_types(i).Height & " Available:" & .bin_types(i).number_available
        Next i
    End With
             
    '''''''''''''''''''''''''''//// Set "Instance" Options; Determine the Max Potential Profit - The "Upper Bound" ///
    result = 0
    ' for mandatory bins and possible items, determine the total possible profit for items and the total cost for bins
    With item_list
        For i = 1 To .num_item_types
            If .item_types(i).mandatory <> -1 Then
                result = result + (.item_types(i).profit * .item_types(i).number_requested)
            End If
        Next i
    End With
    'set max possible profit
    instance.global_upper_bound = result
    '''' Setting Compatability Sheets to False- everything is compatible (for now)
    ' note: This could change if I want to run the program for head and jamb trim at the same time
    instance.item_item_compatibility_worksheet = False
    instance.bin_item_compatibility_worksheet = False
    ' Set Cuts to False
    instance.guillotine_cuts = False
    
    '''''''''''''''''''''''''''//// No compatability data since I'm not using that feature ///
    '''''''''''''''''''''''''''//// Sort Items in Order of Ascending Profit,Importance,Area ///
    Call SortItems
    
    '''''''''''''''''''''''''''//// Start Solution Setup Stuff ///
    Dim incumbent As solution_data
    Call InitializeSolution(incumbent)
        
    Dim best_known As solution_data
    Call InitializeSolution(best_known)
    best_known = incumbent
    
    Dim iteration As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim start_time As Date
    Dim end_time As Date
    Dim continue_flag As Boolean
    
    'infeasibility check
    
    Dim infeasibility_count As Long
    Dim infeasibility_string As String
    
    'checks to see that all items can fit, etc
    'Call FeasibilityCheckData(infeasibility_count, infeasibility_string)

    
    start_time = Timer
    end_time = Timer
        
    'constructive phase
    If solver_options.show_progress = True Then
        Application.ScreenUpdating = True
        Application.StatusBar = "Constructive phase..."
        Application.ScreenUpdating = False
    Else
        Application.ScreenUpdating = True
        Application.StatusBar = "LNS algorithm running..."
        Application.ScreenUpdating = False
    End If
    
    Call SortBins(incumbent)
    For i = 1 To incumbent.num_bins
        For j = 1 To item_list.num_item_types
            continue_flag = True
            Do While (incumbent.unpacked_item_count(incumbent.item_type_order(j)) > 0) And (continue_flag = True)
                continue_flag = AddItemToBin(incumbent, i, incumbent.item_type_order(j), 1)
            Loop
        Next j
        
        incumbent.feasible = True
        For j = 1 To item_list.num_item_types
            If (incumbent.unpacked_item_count(j) > 0) And (item_list.item_types(j).mandatory = 1) Then
                incumbent.feasible = False
                Exit For
            End If
        Next j
        
        Call CalculateDistance(incumbent)

        If ((incumbent.feasible = True) And (best_known.feasible = False)) Or _
           ((incumbent.feasible = False) And (best_known.feasible = False) And (incumbent.total_area > best_known.total_area + epsilon)) Or _
           ((incumbent.feasible = True) And (best_known.feasible = True) And (incumbent.net_profit > best_known.net_profit + epsilon)) Or _
           ((incumbent.feasible = True) And (best_known.feasible = True) And (incumbent.net_profit > best_known.net_profit - epsilon) And (incumbent.total_area < best_known.total_area - epsilon)) Then

            best_known = incumbent
            
        End If
        
    Next i
    
    end_time = Timer
    'MsgBox "Constructive phase result: " & best_known.net_profit & " time: " & end_time - start_time
    
    'improvement phase

    iteration = 0

    Do
        DoEvents
        
        If (solver_options.show_progress = True) And (iteration Mod 100 = 0) Then
            Application.ScreenUpdating = True
            If SteelMode Then
                If best_known.feasible = True Then
                    Application.StatusBar = "Starting " & FOType & " " & InputType & " " & "Iteration #" & iteration '& ". Best net profit found so far: " & best_known.net_profit ' & " TAU: " & best_known.total_area_utilization
                Else
                    Application.StatusBar = "Starting iteration " & iteration & ". Best net profit found so far: N/A"
                End If
            Else
                If best_known.feasible = True Then
                    Application.StatusBar = "Starting " & FOType & " " & InputType & " Trim " & "Iteration #" & iteration '& ". Best net profit found so far: " & best_known.net_profit ' & " TAU: " & best_known.total_area_utilization
                Else
                    Application.StatusBar = "Starting iteration " & iteration & ". Best net profit found so far: N/A"
                End If
            End If
            Application.ScreenUpdating = False
        End If

        If Rnd() < 0.5 Then 'Rnd() < ((end_time - start_time) / solver_options.CPU_time_limit) ^ 2 Then

             incumbent = best_known

        End If
        
        Call PerturbSolution(incumbent)
        
        Call SortBins(incumbent)

        With incumbent
        
            For i = 1 To .num_bins
    
                For j = 1 To item_list.num_item_types
                    
                    continue_flag = True
                    Do While (.unpacked_item_count(.item_type_order(j)) > 0) And (continue_flag = True)
                        continue_flag = AddItemToBin(incumbent, i, .item_type_order(j), 1)
                    Loop
                Next j
    
                .feasible = True
                For j = 1 To item_list.num_item_types
                    If (.unpacked_item_count(j) > 0) And (item_list.item_types(j).mandatory = 1) Then
                        .feasible = False
                        Exit For
                    End If
                Next j
                
                Call CalculateDistance(incumbent)
                
                If ((.feasible = True) And (best_known.feasible = False)) Or _
                   ((.feasible = False) And (best_known.feasible = False) And (.total_area > best_known.total_area + epsilon)) Or _
                   ((.feasible = False) And (best_known.feasible = False) And (.total_area > best_known.total_area - epsilon)) And (.total_distance < best_known.total_distance - epsilon) Or _
                   ((.feasible = True) And (best_known.feasible = True) And (.net_profit > best_known.net_profit + epsilon)) Or _
                   ((.feasible = True) And (best_known.feasible = True) And (.net_profit > best_known.net_profit - epsilon) And (.total_area < best_known.total_area - epsilon)) Or _
                   ((.feasible = True) And (best_known.feasible = True) And (.net_profit > best_known.net_profit - epsilon) And (.total_area < best_known.total_area + epsilon)) And (.total_area_utilization > best_known.total_area_utilization) Or _
                   ((.feasible = True) And (best_known.feasible = True) And (.net_profit > best_known.net_profit - epsilon) And (.total_area < best_known.total_area + epsilon)) And (.total_area_utilization >= best_known.total_area_utilization) And (.total_distance < best_known.total_distance - epsilon) Then
        
                    best_known = incumbent
                End If
    
            Next i
            
        End With

        iteration = iteration + 1
        
        end_time = Timer
        
    Loop While end_time - start_time < solver_options.CPU_time_limit
    
    'MsgBox "Iterations performed: " & iteration
    
BPP_Solver_Finish:
    
    'write the solution
    
    'MsgBox best_known.total_distance
    
    If best_known.feasible = True Then
        Call WriteSolution(best_known, SolvedCollection, InputType)
    ElseIf infeasibility_count > 0 Then
        Call WriteSolution(best_known, SolvedCollection, InputType)
    Else
        'reply = MsgBox("The best found solution after " & iteration & " LNS iterations does not satisfy all constraints. Do you want to overwrite the current solution with the best found solution?", vbYesNo, "BPP Spreadsheet Solver")
        reply = vbYes
        If reply = vbYes Then
            Call WriteSolution(best_known, SolvedCollection, InputType)
        End If
    End If
    
    'Erase the data
    Erase item_list.item_types
    Erase bin_list.bin_types
    Erase compatibility_list.item_to_item
    Erase compatibility_list.bin_to_item

    For i = 1 To incumbent.num_bins
        Erase incumbent.bin(i).items
    Next i
    Erase incumbent.bin
    Erase incumbent.unpacked_item_count
    
    For i = 1 To best_known.num_bins
        Erase best_known.bin(i).items
    Next i
    Erase best_known.bin
    Erase best_known.unpacked_item_count
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    

End Sub




Private Sub SortItems()

    Dim i As Long
    Dim j As Long
    
    Dim swap_item_type As item_type_data
    
    If item_list.num_item_types > 1 Then
       For i = 1 To item_list.num_item_types
           For j = item_list.num_item_types To 2 Step -1
            'checks: 1. if previous item is mandatory and current item isn't isnt
                    '2. If items are equally mandatory AND current item has a sort criteria field value higher than the previous item
                    '3. both not mandatory and current is more profitable than the previous
                    'overall, sub sorts items in ascending importance/length
               If (item_list.item_types(j).mandatory > item_list.item_types(j - 1).mandatory) Or _
                   ((item_list.item_types(j).mandatory = 1) And (item_list.item_types(j - 1).mandatory = 1) And (item_list.item_types(j).sort_criterion > item_list.item_types(j - 1).sort_criterion)) Or _
                   ((item_list.item_types(j).mandatory = 0) And (item_list.item_types(j - 1).mandatory = 0) And ((item_list.item_types(j).profit / item_list.item_types(j).Area) > (item_list.item_types(j - 1).profit / item_list.item_types(j - 1).Area))) Then
                   
                   swap_item_type = item_list.item_types(j)
                   item_list.item_types(j) = item_list.item_types(j - 1)
                   item_list.item_types(j - 1) = swap_item_type
                   
               End If
           Next j
       Next i
    End If
    
'    For i = 1 To item_list.num_item_types
'       MsgBox item_list.item_types(i).id & " " & item_list.item_types(i).width & " " & item_list.item_types(i).height & " "
'    Next i
 
End Sub

Private Sub CalculateDistance(solution As solution_data)

    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim item_flag As Boolean
    Dim bin_count As Long
    Dim penalty As Double 'for not fitting an item type into a single bin
    
    If instance.guillotine_cuts = True Then
    
        With solution
        
            .total_distance = 0
            
            For j = 1 To .num_bins
            
                 With .bin(j)
                     For k = 1 To .item_cnt
                        solution.total_distance = solution.total_distance + .items(k).cut_length
                     Next k
                 End With
    
            Next j
            
        End With
            
    Else
    
        penalty = 1000 ' perhaps find a better value here?
        
        With solution
        
            .total_distance = 0
            
            For j = 1 To .num_bins
            
                 With .bin(j)
                     For k = 1 To .item_cnt
                         For l = k + 1 To .item_cnt
                             If .items(k).item_type = .items(l).item_type Then
                                 solution.total_distance = solution.total_distance + Abs(.items(k).ne_x + .items(k).sw_x - .items(l).ne_x - .items(l).sw_x) + Abs(.items(k).ne_y + .items(k).sw_y - .items(l).ne_y - .items(l).sw_y)
                             End If
                         Next l
                     Next k
                 End With
    
            Next j
            
            For i = 1 To item_list.num_item_types
                bin_count = 0
                For j = 1 To .num_bins
    
                     With .bin(j)
                         item_flag = False
                         For k = 1 To .item_cnt
                                If .items(k).item_type = i Then
                                    item_flag = True
                                    Exit For
                                End If
                         Next k
                         
                         If item_flag = True Then bin_count = bin_count + 1
                     End With
        
                Next j
                
                solution.total_distance = solution.total_distance + penalty * bin_count * bin_count
            Next i
        End With
    End If
    
    With solution
        
        .total_area_utilization = 0
        
        For j = 1 To .num_bins
            .total_area_utilization = .total_area_utilization + ((.bin(j).area_packed / .bin(j).Area) ^ 2)
        Next j
    
    End With
    
End Sub

'' ribbon calls and tab activation
'
'#If Win32 Or Win64 Or (MAC_OFFICE_VERSION >= 15) Then
'
'Sub BPP_Solver_ribbon_call(control As IRibbonControl)
'    Call BPP_Solver
'End Sub
'
'
'#End If


