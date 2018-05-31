Option Explicit

'Much information held in public data type to pass in and out of functions and subs
'Data type to hold raw information about the problem
Public Type Data
     numbInterestPoints As Integer  ' number of attractions to visit
     numbDays As Integer            ' number of days a visitor will spend in Bath
     distances() As Variant         ' distances between POIs
     timeAtPOI() As Long            ' time the tourist plans to spend at each POI
     timeAvailable As Long          ' time available for each day (in minutes)
     interest() As Double           ' score the tourist gives to each POI
     visited() As Boolean           ' mark the visited locations
     interestRatio() As Double      ' holds value for time/interest for each point
     removed() As Boolean           ' for method to delete low interest items from sequence - marks as tried to be deleted, if no benefit, added back in
End Type

'Data type to hold information about solution sequences
Public Type Solution
    score As Integer                ' sum of the scores of the visited POIs
    Feasible As Boolean             ' true if solution feasible, false otherwise
    Sequence() As Integer           ' Sequence in which POIs are visited
    dailyLastPoint() As Long        ' holds the number of the last point visited on each day for the solution
    dailyTimeUsed() As Long         ' holds the time used on each day for the solution
    dailyLastIndex() As Long        ' holds the point in the sequence that the last point in the day happens
    interestScoreInSeq() As Long    ' holds the interest ratio each point in the solution at a given point in the sequence - includes travel time
    

End Type

    
Public Sub TouristTour()
    
    'Set variables to hold times to compare time to run algorithm
    'To do this accurately, all message boxes must be removed/commented out
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    StartTime = Timer
    
    'Set d to hold raw data
    Dim d As Data
    
    'Set some variables to hold summary information about the implementation
    Dim checkSol As Solution
    Dim allFeasible As Boolean
    allFeasible = True
    
    'Indices for later use
    Dim i, j As Long
    
    'Set number of visitors to solve
    Dim numVisitors As Long
    numVisitors = 5
    
    'Activate attractions sheet
    Sheets("Attractions and travelling time").Select
    
    'Set how many POIs there are - calculates dynamically so more interest points could be added
    d.numbInterestPoints = WorksheetFunction.Count(Range("A:A")) - 1
    
    'Read matrix with distances - 43 + 1 for train station
    ReDim d.distances(1 To d.numbInterestPoints + 1, 1 To d.numbInterestPoints + 1) As Variant
    For i = 1 To d.numbInterestPoints + 1
        For j = 1 To d.numbInterestPoints + 1
            d.distances(i, j) = Sheets("Attractions and travelling time").Cells(3 + i, 2 + j)
        Next j
    Next i
    
    'For each visitor do the following
    For i = 1 To numVisitors
    
        'Allocate memory
        ReDim d.interest(0 To d.numbInterestPoints + 1) As Double
        ReDim d.timeAtPOI(0 To d.numbInterestPoints + 1) As Long
        ReDim d.removed(0 To d.numbInterestPoints + 1) As Boolean
        
        'Activate relevant spreadsheet and read time available and number of days from the relevant spreadsheet
        Dim nameSheet As String
        nameSheet = "visitor " & i
        Sheets(nameSheet).Select
        d.timeAvailable = Sheets(nameSheet).Cells(2, 3).Value
        d.numbDays = Sheets(nameSheet).Cells(2, 5).Value
        
        'Read interest and time at POI - add the station as point 1 for consistency
        d.interest(1) = 0
        d.timeAtPOI(1) = 0
        For j = 2 To d.numbInterestPoints + 1
            d.interest(j) = Sheets(nameSheet).Cells(3 + j, 3).Value
            d.timeAtPOI(j) = Sheets(nameSheet).Cells(3 + j, 4).Value
            d.removed(j) = False
        Next j
    
        
        '''Constructive'''
        'Initialise solution for constructive heuristic - sets as maximum length, including all interest points and days
        Dim solConstr As Solution
        ReDim solConstr.Sequence(0 To (d.numbInterestPoints + d.numbDays + 1)) As Integer
        'Set values of -1
        For j = 1 To d.numbInterestPoints + d.numbDays + 1
            solConstr.Sequence(j) = -1
        Next j

        'Call constructive heuristic and evaluate if feasible
        Call Constructive(solConstr, d)
        checkSol = EvaluateSolutionFun(d, solConstr)
        If checkSol.Feasible = False Then
            allFeasible = False
        End If
        
        
        '''Local search'''
        'initialise a solution and data structures
        Dim solLS As Solution
        ReDim solLS.Sequence(0 To (d.numbInterestPoints + d.numbDays + 1)) As Integer
        For j = 1 To d.numbInterestPoints + d.numbDays + 1
            solLS.Sequence(j) = -1
        Next j

        'Call local search and evaluate solution
        Call LS(solLS, d, solConstr)
        checkSol = EvaluateSolutionFun(d, solLS)
        If checkSol.Feasible = False Then
            allFeasible = False
        End If

    Next i
    
    SecondsElapsed = Round(Timer - StartTime, 2)
    MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation
    
    If allFeasible = True Then
        'single message box to conclude - message boxes limited, instead data printed to excel
        'message boxes were present in the code for error checking but have been commented out for attractiveness of final product
        MsgBox "Feasible routes have been found for all visitors. See each visitors' spreadsheet for details."
    End If
    
    
    
End Sub




Public Sub Constructive(sol As Solution, d As Data)
    
    'Set matrix to hold a value to represent the benefit to the solution of moving between any two points
    ReDim d.interestRatio(1 To d.numbInterestPoints + 1, 0 To d.numbInterestPoints + 1)
    
    'set indices for later use
    Dim i, j As Long
    
    ' for each attraction and the train station, set the interest ratio to travel between 2 points (i and j)
    For i = 1 To d.numbInterestPoints + 1
    
        'set large interest ratio for i,0 to act as a base level to improve on later
        d.interestRatio(i, 0) = 1000
        
        For j = 2 To d.numbInterestPoints + 1
            If d.interest(j) <> 0 Then
                'interest ratio = the cost to travel to that destination from each point i, add the time at the attraction all divided by the interest
                d.interestRatio(i, j) = (d.distances(i, j) + d.timeAtPOI(j)) / d.interest(j)
            Else
                'Error handling to cope with 0 values in interest, prevent /0 error
                'if 0 set value as large; 1000 is larger than any time and therefore interest ratio - not overly large as this would be inefficient
                d.interestRatio(i, j) = 1000
            End If
        Next j
        'set i --> i as large and i --> 1 as large to prevent going back to self or to the train station without specific request
        d.interestRatio(i, i) = 10000
        d.interestRatio(i, 1) = 1000
    Next i
                       
    '''''''''''''''Nearest neighbour algorithm'''''''''''''''
    
    'Set starting point as the train station
    Dim startingPoint As Long
    startingPoint = 1
    sol.Sequence(1) = startingPoint
    
    'current is the last city in the solution sequence
    'bestCurrent is the next possible city
    Dim current As Long
    Dim bestCurrent As Long
    
    'initialise all the interest points have not been visited
    ReDim d.visited(1 To (d.numbInterestPoints + 1))
    For i = 1 To d.numbInterestPoints + 1
        d.visited(i) = False
    Next i
    
    'mark the starting point (Bath Spa Station) as visited
    d.visited(1) = True
    
    'dist is a large number which helps to find the minimum interestRatio at the beginning
    'tourTime estimate the possible total time and it will be reset to 0 at the end of below j loop
    Dim dist As Double
    Dim tourTime As Double
    
    'preferences count the interest scores for each visitor and it be initialised to 0
    Dim preferences As Long
    preferences = 0

    'selectedTourTime counts the time that the visitor has already spent
    Dim selectedTourTime As Double
    selectedTourTime = 0
    
    ' m is a counter to detail which day attraction steps are being added to
    ' sequenceCounter1 and SequenceCounter2 define the startingPoint at the beginning of day 2 and day 3
    Dim m As Long
    Dim sequenceCounter1 As Long
    Dim sequenceCounter2 As Long
    Dim endDayNow As Boolean
        
    For m = 1 To d.numbDays
    
        'initialise values at the start of each day
        current = startingPoint
        bestCurrent = 1
        selectedTourTime = 0
        
        'for each point other than the starting point
        For i = 2 To d.numbInterestPoints + 1
            
            'set dist as large
            dist = 1000
            
            'this j loop finds the point with the minimum interest ratio from the previous point
            For j = 2 To d.numbInterestPoints + 1
                tourTime = selectedTourTime + d.distances(current, j) + d.timeAtPOI(j)
                 
                If d.interestRatio(current, j) < dist And d.visited(j) = False And (tourTime + d.distances(bestCurrent, 1)) <= d.timeAvailable Then
                    'best j (distination) attraction is saved as best current, and dist is updated to have the relevant interest ratio
                    bestCurrent = j
                    dist = d.interestRatio(current, j)
                End If
            Next j
            
            'reset/set endDayNow as false; if conditions met day will be ended
            endDayNow = False
            
            'Add the attraction with the best interest ratio to the relevant day
            'added in sequence; sequence counters remember which day it is and -1 and -2 account for the gap to add the home index (1) to the sequence
            If sol.Sequence(i - 1) <> bestCurrent Then
                If m = 1 Then
                    sol.Sequence(i) = bestCurrent
                ElseIf m = 2 Then
                    sol.Sequence(sequenceCounter1 + i - 1) = bestCurrent
                ElseIf m = 3 Then
                    sol.Sequence(sequenceCounter2 + i - 2) = bestCurrent
                End If
            Else
                'if previous solution is the same as the new solution, the day is ended (without 32 is present twice in route 4)
                endDayNow = True
            End If
            
            ' calculate the new total tour time now a point is added
            selectedTourTime = selectedTourTime + d.distances(current, bestCurrent) + d.timeAtPOI(bestCurrent)
                
            'If remaining tour time is bigger than the journey back to location 1 from the suggested solution, replace the value
            ' with 1 and so the day ends. If endDayNow is true, day automatically ends.
            If (selectedTourTime + d.distances(bestCurrent, 1)) > d.timeAvailable And m = 1 Or (m = 1 And endDayNow = True) Then
                sol.Sequence(i) = startingPoint
                'memorize sequenceCounter1 as the last one of the sol.sequence of day 1
                sequenceCounter1 = i
                Exit For
            ElseIf (selectedTourTime + d.distances(bestCurrent, 1)) > d.timeAvailable And m = 2 Or (m = 2 And endDayNow = True) Then
                sol.Sequence(sequenceCounter1 + i - 1) = startingPoint
                'memorize sequenceCounter2 as the last one of the sol.sequence of day 2
                sequenceCounter2 = sequenceCounter1 + i
                Exit For
            ElseIf (selectedTourTime + d.distances(bestCurrent, 1)) > d.timeAvailable And m = 3 Or (m = 3 And endDayNow = True) Then
                sol.Sequence(sequenceCounter2 + i - 2) = startingPoint
                Exit For
            End If
            
            'update global data
            current = bestCurrent
            d.visited(bestCurrent) = True
            
            'calculate total interest score
            preferences = preferences + d.interest(current)
            
        Next i
    Next m

    'MsgBox ("The constructive interest score of this visitor is: " & preferences)

    'output the route onto each spreadsheet
    Dim routeOutputs As Variant
    Cells(4, 8) = "Const Solution"
    For i = 1 To d.numbInterestPoints + d.numbDays + 1
        If sol.Sequence(i) <> -1 Then
            Cells(i + 4, 8) = sol.Sequence(i)
        End If
    Next i
    
    'output the interest scores and times
    Cells(6 + d.numbDays, 11) = "Interest"
    Cells(6 + d.numbDays, 12) = calculateInterest(d, sol)

    Cells(5 + d.numbDays, 11) = "Total time"
    Cells(5 + d.numbDays, 12) = CalculateCost(sol, d)
    
    'update times to split time taken by day then detail by day
    sol = updateTimes(d, sol)
    Cells(4, 12) = "Constructive"
    For i = 1 To d.numbDays
        Cells(i + 4, 11) = "Day " & i
        Cells(i + 4, 12) = sol.dailyTimeUsed(i)
    Next i
    
End Sub

Public Sub LS(sol As Solution, d As Data, solConstr As Solution)
    
    'Dim counter letters for future use
    Dim i, j, k, l, m As Long
    
    'read the solution data from constructive heuristic
    sol = solConstr
    
    'set variable to hold solution length and calculate solution length
    Dim solLength As Long
    For i = 1 To d.numbInterestPoints + d.numbDays + 1
        If sol.Sequence(i) <> -1 Then
            solLength = solLength + 1
        End If
    Next i
    
    '''''''''''''''SWAP implementation'''''''''''''''

    'Dim variable to hold swapped vertex
    Dim memory As Long
    
    'Dim variables for cost before and after each swap
    Dim newCost As Long
    Dim OldCost As Long
    
    'dim variable to hold whether a swap is beneficial
    Dim improvement As Boolean
   
    'dim dummysolution to hold any evaluations of solutions
    Dim dummysolution As Solution
    Dim solMemMinor As Solution
    Dim solMemMajor As Solution
    ReDim sol.interestScoreInSeq(0 To 100) As Long
   
    'Set i as 0 and then Do loop to keep swapping until condition is met (i = ...)
    i = 0
    
    'Run initial swap and update time taken for days
    sol = runSwap(sol, d)
    sol = updateTimes(d, sol)
    
    'Set variable to hold number of new attraction to visit
    Dim newpoint As Long
    Dim lowestInterestPoint As Long
    
    'TRY DELETE - To the simple swap, a loop has been added to remove low interest attractions from the routes and see if any improvement can
    'be made in their absence. This was shown to have a negative impact on the outcome and so is not included. variables and a loop to include this
    'are detailed here:
    ReDim solMemMajor.Sequence(1 To 2) As Integer
    solMemMajor.Sequence(1) = 0

    'TRY DELETE - this includes a loop to cycle multiple times. If tryDelete is False, this m loop only runs once.
    'Change tryDelete to true if you wish to see this run.
    Do Until m = 20
    Dim tryDelete As Boolean
    tryDelete = False
    If tryDelete = True Then
        m = 19
    End If
    
    sol = updateTimes(d, sol)
    
    
    'SWAP IMPLEMENTATION - For each day...
    For j = 1 To d.numbDays
    
        'reset 'newpoint' value. 0 has been used used as d.distances(i,0) and d.interestRatio(0) is always a non-favourable number (large).
        newpoint = 0
        
        'For each attraction (apart from the bus station)
        For i = 2 To (d.numbInterestPoints + 1)
        
            'If the attractions' time at POI is smaller than the total time available over the days
            'd.removed must be false so items deleted cannot be readded - only readded by later reversal
            If ((d.timeAvailable * d.numbDays) - sol.dailyTimeUsed(1) - sol.dailyTimeUsed(2) - sol.dailyTimeUsed(3)) > d.timeAtPOI(i) _
              And d.interest(i) > 0.5 And d.visited(i) = False And d.removed(i) = False Then
                'then set newpoint as that attraction
                newpoint = i
            Else 'else keep newpoint as 0
                newpoint = 0
            End If
            
            'if newpoint is 0 then skip all of the following and try the next i
            If newpoint <> 0 Then
                
                'Set k counter as the length of the solution and save the current solution in "solMemMinor"
                k = solLength
                solMemMinor = sol
                 
                'Relocate-like loop to shift every point 1 along in the solution. This frees up a gap to add a new site in to the tour
                'works sequentially from the last point of the tour (the final 1) back to the last point of the relevant day (j)
                Do Until sol.Sequence(k) = sol.dailyLastPoint(j)
                    sol.Sequence(k + 1) = sol.Sequence(k)
                    k = k - 1
                Loop
                
                'Add newpoint into tour at the end of the relevant day
                sol.Sequence(k + 1) = newpoint
                
                're-run Swap to find most efficient order of assimilating new point into the route
                sol = runSwap(sol, d)
                 
                'evaluate the solution to see if feasible. If feasible sol.Feasible will become true as a result of running this function
                sol = EvaluateSolutionFun(d, sol)
                 
                'if Sol.feasible is false, make sol = solMem (undo all changes), if solution is feasible, update the times, set the new point as visited
                'and add 1 to the number of points in the sequence.
                If sol.Feasible = False Then
                    sol = solMemMinor
                ElseIf sol.Feasible = True Then
                    sol = updateTimes(d, sol)
                    d.visited(newpoint) = True
                    solLength = solLength + 1
                End If
            End If
        'try next point.
        Next i
    'try next day after all unvisited points tried to add to day 1
    Next j
   
'''''''''''''''' TRY DELETE SUB-ALGORITHM '''''''''''''''''''''''
    'try delete sub-algorithm not included as not beneficial to outcome.
    If tryDelete = True Then
        'if first loop, solution is saved as the major (current best) solution
        'if not the first loop, solMemMajor will not equal 0. Therefore if sol is an improvement, solMemMajor will update to the new current best
        'else solution will undo to become solmemmajor and will loop back. This time the point attempted to be removed will be marked and not deleted
        'again
        If solMemMajor.Sequence(1) <> 0 Then
            If calculateInterest(d, sol) >= calculateInterest(d, solMemMajor) Then
                solMemMajor = sol
            Else
                sol = solMemMajor
            End If
        Else
            solMemMajor = sol
        End If
        lowestInterestPoint = 0
        
        For i = 0 To solLength
            If sol.Sequence(i) < 2 Or d.interest(sol.Sequence(i)) = 0 Then
                sol.interestScoreInSeq(i) = 1000
            Else
                sol.interestScoreInSeq(i) = (d.distances(sol.Sequence(i - 1), sol.Sequence(i)) + d.timeAtPOI(sol.Sequence(i))) / d.interest(sol.Sequence(i))
            End If
            
            If sol.interestScoreInSeq(i) <= sol.interestScoreInSeq(lowestInterestPoint) And d.removed(i) = False And sol.Sequence(i) > 1 Then
                lowestInterestPoint = i
            End If
        Next i
        
        If lowestInterestPoint <> 0 Then
            d.removed(lowestInterestPoint) = True
            
            Do Until sol.Sequence(lowestInterestPoint - 1) = -1
                sol.Sequence(lowestInterestPoint) = sol.Sequence(lowestInterestPoint + 1)
                lowestInterestPoint = lowestInterestPoint + 1
            Loop
            sol = updateTimes(d, sol)
        End If
        
        
    End If
    'end of loop for try delete
    m = m + 1
    Loop
''''''''''''''''''''' END OF TRY DELETE '''''''''''''''''''''''
      
   
    'print final route to relevant spreadsheet
    Dim routeOutputs As Variant
    Cells(4, 9) = "LS Solution"
    For i = 1 To d.numbInterestPoints + d.numbDays + 1
        If sol.Sequence(i) <> -1 Then
            Cells(i + 4, 1 + 8) = sol.Sequence(i)
        End If
    Next i
    
    'print interest and total time to spreadsheet
    Cells(4, 13) = "Local search"
    Cells(6 + d.numbDays, 13) = calculateInterest(d, sol)
    Cells(5 + d.numbDays, 13) = CalculateCost(sol, d)
    Cells(7 + d.numbDays, 11) = "Interest/minute"
    Cells(7 + d.numbDays, 12) = Cells(6 + d.numbDays, 12) / Cells(5 + d.numbDays, 12)
    Cells(7 + d.numbDays, 13) = Cells(6 + d.numbDays, 13) / Cells(5 + d.numbDays, 13)
    
    For i = 1 To d.numbDays
        Cells(i + 4, 13) = sol.dailyTimeUsed(i)
    Next i
    
    

End Sub

'' Function to run swap algorithm and relocate set points within a route; keeping the home point in the same place
Public Function runSwap(sol As Solution, d As Data) As Solution

        Dim memory, i, j, k As Long
        Dim newCost, OldCost As Long
        Dim improvement As Boolean
        Dim dummysolution As Solution
        
        'loop to do this a few times as a change to add, for instance point 25, may make adding point 15 feasible so we must repeat the loop
        'this could be looped until there is no change of for a set time/number of rounds
        i = 1
        
        'set current cost
        OldCost = CalculateCost(sol, d)
        
        Do Until i = 3
            For j = 1 To UBound(sol.Sequence())
                
                'set base values for variables
                memory = sol.Sequence(j)
                k = 1
                improvement = False
                
                'loop to swap vertex j with each other vertex (k) until either all vertices have been tried
                'or there is and improvement. If there is an improvement, leave loop and start from the next j
                'this approach causes diversification by making the swap as soon as its found instead of fully assessing the neighbourhood
                Do Until k = UBound(sol.Sequence()) Or improvement = True
                    
                    'if stops vertices from swapping with itself (when j=l) and only swaps values that are bigger than one. This stops
                    'the swapping of both the train station (1) and the solution list includes -1 after list - stops both being tried.
                    ' ensuring k > j saves code from checking swaps it has already done (it will only swap forward).
                    If j <> k And sol.Sequence(j) > 1 And sol.Sequence(k) > 1 And k > j Then
                        
                        'set j as k and k as j - which was stored in memory previously
                        sol.Sequence(j) = sol.Sequence(k)
                        sol.Sequence(k) = memory
                        
                        'calculate new cost and whether new solution is feasible
                        newCost = CalculateCost(sol, d)
                        sol = EvaluateSolutionFun(d, sol) 'will turn sol.feasible to True or False
                        
                        'if improvement and feasible, set improvement if not, undo change to sequence
                        If newCost < OldCost And sol.Feasible = True Then
                            improvement = True
                            i = i - 1
                            OldCost = newCost
                        Else
                            sol.Sequence(k) = sol.Sequence(j)
                            sol.Sequence(j) = memory
                        End If
                    End If
                      
                'add 1 to counter and loop back to try next k
                k = k + 1
                Loop
            Next j
        i = i + 1
        Loop
        
        'Return updated solution
        runSwap = sol

End Function
Public Function CalculateCost(sol As Solution, d As Data) As Long

    'set initial variables
    Dim pointCost As Long
    CalculateCost = 0
    pointCost = 0
    Dim i As Long
    
    'for each attraction point
    For i = 1 To d.numbInterestPoints + 1
        
        'calculate time to travel to point i and time spent at point i
        If sol.Sequence(i) <> -1 And sol.Sequence(i + 1) <> -1 Then
            pointCost = d.distances(sol.Sequence(i), sol.Sequence(i + 1)) + d.timeAtPOI(sol.Sequence(i))
            
            'if next point is -1 (is after end of solution) exit for
        ElseIf sol.Sequence(i + 1) = -1 Then
            Exit For
        End If
        
        'Update calculate cost value
        CalculateCost = CalculateCost + pointCost
    Next i
    
    'CalculateCost is returned

End Function
Public Function updateTimes(d As Data, sol As Solution) As Solution
    
    'set indices for future use
    Dim i, j, k As Long
    
    'updatetimes function starts as solution fed into the function
    updateTimes = sol
    
    'set initial variables
    j = 1
    ReDim updateTimes.dailyTimeUsed(1 To 3)
    ReDim updateTimes.dailyLastPoint(1 To 3)
    ReDim updateTimes.dailyLastIndex(1 To 3)
    
    'for each day
    For i = 1 To d.numbDays
    
        'if point j is not -1 (at end of sequence it is -1)
        If sol.Sequence(j) <> -1 Then
            
            'add the journey from the previous point to the new point and the time spent at the new point
            'do this until the sequence value = 1 (the train station) which signals a new day.
            ' this gives you the time used on each day as well as the name of the last point of each day and its point in the sequence.
            Do
                j = j + 1
                updateTimes.dailyTimeUsed(i) = updateTimes.dailyTimeUsed(i) + d.timeAtPOI(sol.Sequence(j)) + d.distances(sol.Sequence(j - 1), sol.Sequence(j))
                
            Loop Until sol.Sequence(j) = 1
            updateTimes.dailyLastPoint(i) = sol.Sequence(j - 1)
            updateTimes.dailyLastIndex(i) = j - 1
        End If
    Next i
       
    
End Function
Public Function calculateInterest(d As Data, sol As Solution) As Long
    
    'set initial variables
    Dim i As Long
    calculateInterest = 0
    
    'like calculateCost, work through each point summing interest
    Do Until sol.Sequence(i) = -1
        calculateInterest = calculateInterest + d.interest(sol.Sequence(i))
        i = i + 1
    Loop
    
End Function

'Both a function for evaluating the solution and a sub are included.
'The function is more streamlined and used within the algorithm to check for feasible improvements.
'The sub is a more detailed, complete check of the feasibility; it is only ran once so can be in depth.
Public Function EvaluateSolutionFun(d As Data, sol As Solution) As Solution
    
    'set initial variables
    Dim i, j As Integer
    Dim Duration As Integer
    
    'function starts as solution fed in; function works to disprove statement .feasible = True'
    EvaluateSolutionFun = sol
    EvaluateSolutionFun.Feasible = True
    
    'now we start checking the feasibility of the solution
    'we first check that the solution contains the correct number of days of visit
    Dim countDays As Integer
    
    'count days
    For i = 1 To d.numbInterestPoints + d.numbDays + 1
        If sol.Sequence(i) = 1 Then countDays = countDays + 1
    Next i
      
    'if days in solution isn't the same as that prescribed, function fails and jumps to end of function
    If (countDays <> d.numbDays + 1) Then
        EvaluateSolutionFun.Feasible = False
        GoTo jmp
    End If
    
    
    j = 1
    
    'for each day check the time limit isn't breached
    For i = 1 To d.numbDays
        
        'initially set duration as 0
        Duration = 0
        
        'for each step that isn't bigger than one calculate the daily time
        Do While sol.Sequence(j + 1) > 1
            'we add up the travelling durations up to the current POI (including the travelling time to the POI and the visiting time)
            Duration = Duration + d.timeAtPOI(sol.Sequence(j + 1)) + d.distances(sol.Sequence(j), sol.Sequence(j + 1))
            j = j + 1
        Loop
        
        ' if duration + duration to last point is more than the time available, function fails and jumps to end
        Duration = Duration + d.distances(sol.Sequence(j), 1)
        If Duration > d.timeAvailable Then
            EvaluateSolutionFun.Feasible = False
            GoTo jmp
        End If
        j = j + 1
    Next i
     
jmp:
End Function

