VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "INotifyPropertyChanged"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents an observer pattern subject that can notify observers when a property value changes."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Represents an observer pattern subject that can notify observers when a property value changes. </summary>
''' <remarks>
''' This interface is central in the event propagation mechanics: in design patterns terms, a class that
''' implement it is the subject in an Observer Pattern where the registered handlers are the observers.
''' The reason a ViewModel needs to implement this interface, is because creating a property binding
''' registers the binding as an observer – and it handles the ViewModel telling it about a property change
''' by applying the binding(s) for that property.
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Notifies all registered handlers of a property value change. </summary>
''' <remarks>
''' The <c>OnPropertyChanged</c> method is only invoked when the assigned property value is
''' different than the current value, and we don’t need to worry about tracking/adding observers
''' or invoking them, because everything we need is already encapsulated in the
''' <see cref="PropertyChangeNotifierBase"/> object, so we implement the interface by simply passing the
''' parameters over to this “notifier” instance.
''' </remarks>
''' <param name="a_eventArgs">   An <see cref="cc_isr_Core_IO.PropertyChangedEventArgs"/> event args. </param>
Public Event PropertyChanged(ByVal a_eventArgs As cc_isr_Core.PropertyChangedEventArgs)
