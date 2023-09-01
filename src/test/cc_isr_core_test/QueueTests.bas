Attribute VB_Name = "QueueTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests. Queue methods. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   [Unit Test]. Tests Constructing the Queue. </summary>
Public Function TestQueueShouldConstruct() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Dim p_expectedCount As Integer:    p_expectedCount = 0
    Dim p_expectedCapacity As Integer: p_expectedCapacity = 11
    Dim p_queue As Queue: Set p_queue = cc_isr_Core_IO.Factory.CreateQueue(p_expectedCapacity)
    
    Set p_outcome = cc_isr_Test_Fx.Assert.IsNotNothing(p_queue, _
                                    "A queue should be created.")

    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCount, _
                                        p_queue.Count, _
                                        "The Queue count should initialize at " & VBA.CStr(p_expectedCount) & ".")
    End If

    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCapacity, _
                                        p_queue.Capacity, _
                                        "The Queue should have the expected capacity.")
    End If

    Debug.Print p_outcome.BuildReport("TestQueueShouldConstruct")
    
    Set TestQueueShouldConstruct = p_outcome

End Function

''' <summary>   [Unit Test]. The Queue should enqueue, dequeue and peek. </summary>
Public Function TestQueueShouldEnqueue() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Dim p_expectedCount As Integer:    p_expectedCount = 0
    Dim p_expectedCapacity As Integer: p_expectedCapacity = 11
    Dim p_queue As Queue: Set p_queue = cc_isr_Core_IO.Factory.CreateQueue(p_expectedCapacity)
    
    Set p_outcome = cc_isr_Test_Fx.Assert.IsNotNothing(p_queue, _
                                    "A queue should be created.")

    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCount, _
                                        p_queue.Count, _
                                        "The Queue count should initialize at " & VBA.CStr(p_expectedCount) & ".")
    End If

    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCapacity, _
                                        p_queue.Capacity, _
                                        "The Queue should have the expected capacity.")
    End If

    Dim p_lastItem As Integer
    
    If p_outcome.AssertSuccessful Then
        p_lastItem = 1
        p_expectedCount = p_expectedCount + 1
        p_queue.Enqueue VBA.CStr(p_lastItem)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCount, _
                                        p_queue.Count, _
                                        "The Queue count should increment after enqueueing.")
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Dim i As Integer
        For i = 1 To p_queue.Capacity - 1
            p_lastItem = p_lastItem + 1
            p_queue.Enqueue VBA.CStr(p_lastItem)
            DoEvents
        Next i
        p_expectedCount = p_queue.Capacity
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCount, _
                            p_queue.Count, _
                            "The Queue count should set to full count after adding " & VBA.CStr(p_queue.Capacity) & " items.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_lastItem = p_lastItem + 1
        p_queue.Enqueue VBA.CStr(p_lastItem)
        p_expectedCount = p_queue.Capacity
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCount, _
                            p_queue.Count, _
                            "The Queue count should remain at full capacity after adding beyond the queue capacity.")
    End If
    
    Dim p_index As Integer
    Dim p_expectedValue As Integer
    Dim p_actualValue As String
    If p_outcome.AssertSuccessful Then
        p_index = 1
        p_expectedValue = IIf(p_lastItem > p_queue.Capacity, p_lastItem - p_queue.Capacity + 1, p_index)
        p_actualValue = p_queue.Peek(p_index)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(VBA.CStr(p_expectedValue), _
                            p_actualValue, _
                            "The Queue should have the expected value at the " & VBA.CStr(p_index) & " index.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_index = p_queue.Capacity
        p_expectedValue = IIf(p_lastItem > p_queue.Capacity, p_lastItem, p_index)
        p_actualValue = p_queue.Peek(p_index)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(VBA.CStr(p_expectedValue), _
                            p_actualValue, _
                            "The Queue should have the expected value at the " & VBA.CStr(p_index) & " index.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_index = 1
        p_expectedValue = IIf(p_lastItem > p_queue.Capacity, p_lastItem - p_queue.Capacity + 1, p_index)
        While p_queue.Count > 0
            DoEvents
            p_actualValue = p_queue.Dequeue()
            If p_outcome.AssertSuccessful Then
                Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(VBA.CStr(p_expectedValue), _
                                    p_actualValue, _
                                    "The Queue should have the expected value at the " & VBA.CStr(p_index) & " index after dequeuing to " & p_queue.Count & " items.")
                p_expectedValue = p_expectedValue + 1
                p_index = p_index + 1
            End If
        Wend
        
    End If
    
    
    Debug.Print p_outcome.BuildReport("TestQueueShouldEnqueue")
    
    Set TestQueueShouldEnqueue = p_outcome

End Function



