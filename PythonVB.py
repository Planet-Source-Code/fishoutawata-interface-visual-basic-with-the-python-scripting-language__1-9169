class PythonVB:
    #Public Methods available to VB    
    _public_methods_ = ['SayHello']
    #Class Name Created in VB.
    #Example Set objPython = CreateObject("PythonVB.Demo")
    _reg_progid_ = "PythonVB.Demo"
    #Never Copy this GUID!  Use pythoncom module to create a new one
    #Example: import pythoncom
    #         print pythoncom.CreateGuid() 
    _reg_clsid_ = "{9C077AB6-0611-4A5C-8628-C78CD96018EF}"

    def SayHello(self):
        return "Hello from Python!!"

if __name__ =='__main__':
    print "Registering COM Server..."
    import win32com.server.register
    win32com.server.register.UseCommandLine(PythonVB)
    
    
