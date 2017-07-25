import uno
import unohelper

#gStyle = 'Heading 3'


class OfficeHelper:
    '''Frequently used methods in office context'''

    def __init__(self, ctx=uno.getComponentContext()):
        self.ctx = ctx
        #self.smgr = self.ctx.ServiceManager
        self.resolver = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", self.ctx)
        self.context = self.resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        self.smgr = self.context.ServiceManager
        self.dispatcher = self.smgr.createInstance("com.sun.star.frame.DispatchHelper")

    def convertToURL(self,sPath):
        return unohelper.systemPathToFileUrl(sPath)

    def convertFromURL(self,sURL):
        return unohelper.fileUrlToSystemPath(sURL)

    def copyCurrentSelection(self):
        return self.dispatcher.executeDispatch(oHelper.getCurrentFrame(), ".uno:Copy","",0,())

    def pasteCurrentSelection(self):
        return self.dispatcher.executeDispatch(oHelper.getCurrentFrame(), ".uno:Paste","",0,())

    def createNewScalc(self):
        return self.getDesktop().loadComponentFromURL("private:factory/scalc","_blank",0,())

    def importCSV(self,sCSVFileURL):
        from com.sun.star.beans import PropertyValue
        tDataProperties = []
        p = PropertyValue()

        p.value.Name = "FilterName"
        p.value.Value = "Text - txt - csv (StarCalc)"
        tDataProperties.append(p)

        p.value.Name = "FilterOptions"
        p.value.Value = "59,34,IBMPC_850,1,,0,false,true"
        tDataProperties.append(p)

        return self.getDesktop().loadComponentFromURL(sCSVFileURL,"_blank",0,tDataProperties)

    def createUnoService(self, service):
        return self.smgr.createInstance(service)

    def getDesktop(self):
        return self.smgr.createInstanceWithContext("com.sun.star.frame.Desktop", self.context)

    def getCurrentComponent(self):
        return self.getDesktop().getCurrentComponent()

    def getCurrentFrame(self):
        return self.getDesktop().getCurrentFrame()

    def getCurrentComponentWindow(self):
        return self.getCurrentFrame().getComponentWindow()

    def getCurrentContainerWindow(self):
        return self.getCurrentFrame().getContainerWindow()

    def getCurrentController(self):
        return self.getCurrentFrame().getController()

    def callMRI(self, obj=None):
        '''Create an instance of MRI inspector and inspect the given object (default is selection)'''
        if not obj:
            obj = self.getCurrentController().getSelection()
        mri = self.createUnoService("mytools.Mri")
        mri.inspect(obj)


'''def cleanParaStyle():
    ohelper = OfficeHelper()
    view = ohelper.getCurrentController()
    frame = view.getFrame()
    doc = view.getModel()
    dispatcher = ohelper.createUnoService("com.sun.star.frame.DispatchHelper")
    oFind = findParaStyle(doc, gStyle)
    view.select(oFind)
    dispatcher.executeDispatch(frame, ".uno:ResetAttributes", "", 0, ())


def findParaStyle(obj, s):
    oDesc = obj.createSearchDescriptor()
    oDesc.SearchAll = True
    oDesc.setSearchString(s)
    oDesc.SearchStyles = True
    return obj.findAll(oDesc)'''