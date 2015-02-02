# Continuously updating word count based on paragraph style
import unohelper, uno, time
from com.sun.star.awt import XTopWindowListener
import threading

DEBUG = False
#DEBUG = True
def PrintOut(st):
    if DEBUG: XSCRIPTCONTEXT.getDocument().Text.End.String = '\n' + st

def getWordCountTarget():
    retval = 0

    # Only if the field exists
    if doc.getTextFieldMasters().hasByName('com.sun.star.text.FieldMaster.User.WordCountGoal'):
        # Get the field
        wordcountTarget = doc.getTextFieldMasters().getByName('com.sun.star.text.FieldMaster.User.WordCountTarget')
        retval = int(wordcountTarget.Content)

    return retval

def setWordCountTarget(target):
    if doc.getTextFieldMasters().hasByName('com.sun.star.text.FieldMaster.User.WordCountTarget'):
        wordcountTarget = doc.getTextFieldMasters().getByName('com.sun.star.text.FieldMaster.User.WordCountTarget')
    else:
        wordcountTarget = doc.createInstance('com.sun.star.text.FieldMaster.User')
        wordcountTarget.Name = 'WordCountTarget'

    wordcountTarget.Content = target
    # Refresh the field if inserted in the document from Insert > Fields >
    # Other... > Variables > Userdefined fields
    doc.TextFields.refresh()

def getWordCountStyles():
    retval = 'All'

    # Only if the field exists
    if doc.getTextFieldMasters().hasByName('com.sun.star.text.FieldMaster.User.WordCountStyles'):
        # Get the field
        wordcountstyles = doc.getTextFieldMasters().getByName('com.sun.star.text.FieldMaster.User.WordCountStyles')
        retval = wordcountstyles.Content

    return retval

def setWordCountStyles(styles):
    if doc.getTextFieldMasters().hasByName('com.sun.star.text.FieldMaster.User.WordCountStyles'):
        wordcountstyles = doc.getTextFieldMasters().getByName('com.sun.star.text.FieldMaster.User.WordCountStyles')
    else:
        wordcountstyles = doc.createInstance('com.sun.star.text.FieldMaster.User')
        wordcountstyles.Name = 'WordCountStyles'

    wordcountstyles.Content = styles
    # Refresh the field if inserted in the document from Insert > Fields >
    # Other... > Variables > Userdefined fields
    doc.TextFields.refresh()

def RGB( nRed, nGreen, nBlue ):
    """Return an integer which repsents a color.
    The color is specified in RGB notation.
    Each of nRed, nGreen and nBlue must be a number from 0 to 255.
    """
    return (int( nRed ) & 255) << 16 | (int( nGreen ) & 255) << 8 | (int( nBlue ) & 255) 

def updateCount(wordCountModel, ProgressBarModel):
    '''Updates the GUI.
    Updates the word count and the percentage completed in the GUI.'''
    global wordcountSnapshot

    try:
        wordcountCurrent = doc.WordCount #model.WordCount
    except AttributeError: return

    if styles in ('All', 'all', 'ALL') or viewCursor.ParaStyleName in styles: #('MS Text', 'MS Text 1st Paragraph of Scene'):
        tempcount = wordcountCurrent - wordcountSnapshot
        if tempcount != 0:
            wordcount = int(wordCountModel.Label) + tempcount
            #wordCountModel.Label = str(wordcount)
            percent = 98 * float(wordcount) / target if target else 0

            if percent > 98:
                percent = 98
            elif percent < 0:
                percent = 0
            R = G = B = 0

            if percent < 75:
                R = 255
                G = (255 * percent) / 98
            elif percent >= 75:
                R = (255 * (75 - percent)) / 75
                G = (255 * percent) / 98

            wordCountModel.Label = str(wordcount)
            ProgressBarModel.Width = percent
            ProgressBarModel.BackgroundColor = RGB(R, G, B)

    wordcountSnapshot = wordcountCurrent

# This is the user interface:

#######################################
# Session Word Count with Style _ o x #
#######################################
#            _____                    #
#     301 /  |500|                    #
#            -----                    #
# ___________________________         #
# |##############           |         #
# ---------------------------         #
#                                     #
# Only include these styles:          #
# ___________________________         #
# |Standard,Heading 1       |         #
# ---------------------------         #
# (Default is All; separate by comma) #
#######################################

# The boxed '500' is a text entry box.
# The boxed 'Standard,Heading 1' is a text entry box. 

class WindowListener(unohelper.Base, XTopWindowListener):
    def __init__(self, workerThread, targetModel, stylesModel, exiting):
        self.workerThread = workerThread
        self.targetModel = targetModel
        self.stylesModel = stylesModel
        self.exiting = exiting

    def windowDeactivated(self, e):
        # Save the target as a text field in the current document
        global target
        try: target = int(self.targetModel.Text)
        except: target = 0

        if target != getWordCountTarget():
            setWordCountTarget(target)

        # Save the styles as a text field in the current document
        global styles
        try: styles = self.stylesModel.Text
        except: styles = 'All'
        if styles == '':
            styles = 'All'
            self.stylesModel.Text = styles
        if styles != getWordCountStyles():
            setWordCountStyles(styles)

    def windowClosing(self, e):
        self.exiting.set()
        # Wait for updater thread to finish before closing window
        self.workerThread.join()
        e.Source.setVisible(False)

def addControl(controlType, dlgModel, x, y, width, height, label, name = None):
    control = dlgModel.createInstance(controlType)
    control.PositionX = x
    control.PositionY = y
    control.Width = width
    control.Height = height
    if controlType == 'com.sun.star.awt.UnoControlFixedTextModel':
        control.Label = label
    elif controlType == 'com.sun.star.awt.UnoControlEditModel':
        control.Text = label
    elif controlType == 'com.sun.star.awt.UnoControlProgressBarModel':
        control.ProgressValue = label

    control.Name = name if name else 'unnamed'
    dlgModel.insertByName(control.Name, control)
    return control

def loopTheLoop(exiting, wordCountModel, ProgressBarModel):
    while not exiting.isSet():
        updateCount(wordCountModel, ProgressBarModel)
        time.sleep(.25)

def wordCountbyStyle(arg = None):
    '''Displays a continuously updating word count.'''
    global viewCursor, doc, wordcountSnapshot, target, styles

    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    #desktop = smgr.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
    #model = desktop.getCurrentComponent()
    #viewCursor = model.getCurrentController().getViewCursor()
    doc = XSCRIPTCONTEXT.getDocument()
    viewCursor = doc.getCurrentController().getViewCursor()

    dialogModel = smgr.createInstanceWithContext('com.sun.star.awt.UnoControlDialogModel', ctx)
    dialogModel.PositionX = XSCRIPTCONTEXT.getDocument().CurrentController.Frame.ContainerWindow.PosSize.Width / 2.2 - 115
    dialogModel.Width = 110
    dialogModel.Height = 70
    dialogModel.Title = 'Session Word Count with Style'

    lblWc = addControl('com.sun.star.awt.UnoControlFixedTextModel', dialogModel, 6, 3, 25, 14, '', 'lblWc')
    lblWc.Align = 2 # Align right
    lblWc.Label = '0'
    addControl('com.sun.star.awt.UnoControlFixedTextModel', dialogModel, 33, 3, 10, 14, ' / ')
    txtTarget = addControl('com.sun.star.awt.UnoControlEditModel', dialogModel, 43, 1, 25, 12, '', 'txtTarget')
    target = getWordCountTarget()
    txtTarget.Text = target

    #ProgressBar = addControl('com.sun.star.awt.UnoControlProgressBarModel', dialogModel, 6, 15, 98, 10, '', 'lblPercent')
    #ProgressBar.ProgressValueMin = 0
    #ProgressBar.ProgressValueMax = 100

    lblProgressBarForeground = addControl('com.sun.star.awt.UnoControlFixedTextModel', dialogModel, 6, 17, 0, 14, '', 'lblProgressBarForeground')
    lblProgressBarBackground = addControl('com.sun.star.awt.UnoControlFixedTextModel', dialogModel, 6, 17, 98, 14, '', 'lblProgressBarBackground')
    lblProgressBarBackground.BackgroundColor = RGB(210,210,210)

    lblStyles1 = addControl('com.sun.star.awt.UnoControlFixedTextModel', dialogModel, 6, 35, 90, 14, '', 'lblStyles1')
    lblStyles1.Label = 'Only include these styles:'
    txtStyles = addControl('com.sun.star.awt.UnoControlEditModel', dialogModel, 6, 45, 98, 14, '', 'txtStyles')
    styles = getWordCountStyles()
    if styles == '':
        styles = 'All'
    txtStyles.Text = styles
    lblStyles2 = addControl('com.sun.star.awt.UnoControlFixedTextModel', dialogModel, 6, 60, 90, 14, '', 'lblStyles2')
    lblStyles2.Label = '(Default is All; separate by comma)'

    controlContainer = smgr.createInstanceWithContext('com.sun.star.awt.UnoControlDialog', ctx)
    controlContainer.setModel(dialogModel)
    controlContainer.setVisible(True)

    targetModel = controlContainer.getControl('txtTarget').getModel()
    wordCountModel = controlContainer.getControl('lblWc').getModel()
    #percentModel = controlContainer.getControl('lblPercent').getModel()
    #ProgressBar.ProgressValue = percentModel.ProgressValue
    stylesModel = controlContainer.getControl('txtStyles').getModel()
    ProgressBarModel = controlContainer.getControl('lblProgressBarForeground').getModel()

    #try:
        #if not model.supportsService('com.sun.star.text.TextDocument'):
            #return
    #except AttributeError:
        #return

    wordcountSnapshot = doc.WordCount #model.WordCount

    exiting = threading.Event()
    updaterThread = threading.Thread(
        target = loopTheLoop
      , args = (exiting, wordCountModel, ProgressBarModel)
    )
    controlContainer.addTopWindowListener(
        WindowListener(updaterThread, targetModel, stylesModel, exiting)
    )
    updaterThread.start()

g_exportedScripts = wordCountbyStyle,

# Standard disclaimer and MIT licence .
# These macros are copyright (c) 2015 Executie.
# Permission is hereby granted, free of charge, to any person
# obtaining a copy of this software and associated documentation
# files (the 'Software'), to deal in the Software without
# restriction, including without limitation the rights to use, copy,
# modify, merge, publish, distribute, sublicense, and/or sell copies
# of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be
# included in all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND,
# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
# MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
# NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
# HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
# WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
# DEALINGS IN THE SOFTWARE.
