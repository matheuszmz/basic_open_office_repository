public function Posicionar (celula as string)
	document = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	
	dim args(0) as new com.sun.star.beans.PropertyValue
	
	args(0).Name = "ToPoint"
	args(0).Value = celula
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args())
end function

public function Digitar (texto as string)
	document = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	
	dim args(0) as new com.sun.star.beans.PropertyValue
	
	args(0).Name = "StringName"
	args(0).Value = texto
	dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args())
end function


public function Copiar (celula as string)
	document = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	
	dim args(0) as new com.sun.star.beans.PropertyValue
	
	args(0).Name = "ToPoint"
	args(0).Value = celula
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args())
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
end function


public function Colar (celula as string)
	document = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	
	dim args(0) as new com.sun.star.beans.PropertyValue
	
	args(0).Name = "ToPoint"
	args(0).Value = celula
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args())
	dispatcher.executeDispatch(document, ".uno:Paste", "", 0, Array())
end function


sub teste
	Posicionar("A1")
  Digitar("Hello World")
	Copiar("A1")
	Colar("B1")
end sub
