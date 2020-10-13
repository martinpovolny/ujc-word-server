from xmlrpc.server import SimpleXMLRPCServer
from xmlrpc.server import SimpleXMLRPCRequestHandler

from time import sleep
import win32com.client as win32

class WordProxy:
    def __init__(self):
        self.docs = {}


    def setup_locals(self, doc_id):
        #self.text   = self.docs[doc_id]['text']
        self.doc    = self.docs[doc_id]['doc']
        self.rng    = self.docs[doc_id]['rng']
        #self.cursor = self.docs[doc_id]['cursor']


    def create_document(self):
        word = win32.gencache.EnsureDispatch('Word.Application')
        doc = word.Documents.Add()
        word.Visible = True
        #sleep(1)
        
        # shortcuts
        #self.text = self.doc.Text
        #self.cursor = self.text.createTextCursor()

        doc_id = uuid.uuid1().hex
        self.docs[doc_id] = {
            #'text':   self.text,
            #'cursor': self.cursor,
            'rng': doc.Range(0,0),
            'doc': doc
        }

        return doc_id


    def set_char_weight(self, doc, weight):
        self.setup_locals(doc)
        if weight == 'NORMAL':
            #w = NORMAL
            self.rng.Bold = False
        if weight == 'BOLD':
            #w = BOLD
            self.rng.Bold = True
        #self.cursor.setPropertyValue ( "CharWeight", w )
        return 1


    def set_char_posture(self, doc, slant):
        self.setup_locals(doc)
        s = NONE
        if slant == 'ITALIC':
            s = ITALIC
        self.cursor.setPropertyValue ( "CharPosture", s )
        return 1


    def set_char_underline(self, doc, underline):
        self.setup_locals(doc)
        u = 0
        if underline == 'SINGLE':
            u = SINGLE
        self.cursor.setPropertyValue ( "CharUnderline", u )
        return 1


    def set_char_height(self, doc, height):
        self.setup_locals(doc)
        self.cursor.setPropertyValue( "CharHeight", height )
        return 1


    def set_char_font_name(self, doc, font_name):
        self.setup_locals(doc)
        self.cursor.setPropertyValue( "CharFontName", font_name )
        return 1


    def insert_control_character(self, doc, char):
        self.setup_locals(doc)
        #if char == 'HARD_SPACE':
            #c = HARD_SPACE
        if char == 'PARAGRAPH_BREAK':
            #c = PARAGRAPH_BREAK
            rng.InsertBreak( win32.constants.wdLineBreak )
        #if char == 'SOFT_HYPHEN':
            #c = SOFT_HYPHEN
        #self.text.insertControlCharacter( self.cursor, c, 0 )
        return 1


    def set_char_escapement(self, doc, size, offset):
        self.setup_locals(doc)
        self.cursor.setPropertyValue( "CharEscapement", size )
        self.cursor.setPropertyValue( "CharEscapementHeight", offset )
        return 1


    def put_text( self, doc, text ):
        self.setup_locals(doc)
        #self.text.insertString( self.cursor, text, 0 );

        #self.rng.Collapse( win32.constants.wdCollapseEnd )
        self.rng.Text = text
        self.rng.Collapse( win32.constants.wdCollapseEnd )

        return 1


    def save_and_close(self, doc, outputfile):
        self.setup_locals(doc)
        #cwd = systemPathToFileUrl( getcwd() )
        #args = ( makePropertyValue("FilterName","MS Word 97"), )
        #destFile = absolutize( cwd, systemPathToFileUrl(outputfile) )
        #self.doc.storeAsURL(destFile, args)

        #try:
        #    self.doc.dispose()
        #except:
        #    print("error while saving doc")

        #RANGE = range(3, 8)
        #def word():
        #    word = win32.gencache.EnsureDispatch('Word.Application')
        #    doc = word.Documents.Add()
        #    word.Visible = True
        #    sleep(1)
        #    rng = doc.Range(0,0)
        #    rng.InsertAfter('Hacking Word with Python\r\n\r\n')
        #    sleep(1)
        #    for i in RANGE:
        #        rng.InsertAfter('Line %d\r\n' % i)
        #        sleep(1)
        #    rng.InsertAfter("\r\nPython rules!\r\n")
        doc.Close(False)
        # word.Application.Quit()

        del self.docs[doc]
        return 1


    def set_para_style(self, doc, style):
        self.setup_locals(doc)
        if (style == 'Default' or style == 'Normal'):
          style = 'Standard'

        self.cursor.ParaStyleName = style

        # #print style
        # #print self.cursor.ParaStyleName
        # #print self.cursor.getPropertyValue( 'ParaStyleName' )

        # style2 = self.cursor.getPropertyValue( 'ParaStyleName' )
        # self.text.insertString( self.cursor, style2, 0 );

        # #print "OK"
        # #self.cursor.setPropertyValue( 'ParaStyleName', style2 )
        #self.cursor.setPropertyValue( 'ParaStyleName', style )

        # style2 = self.cursor.getPropertyValue( 'ParaStyleName' )
        # self.text.insertString( self.cursor, style2, 0 );
        # #self.cursor.setPropertyValue( 'CharStyleName', style )
        return 1


    def set_para_adjust(self, doc, style):
        self.setup_locals(doc)
        if style == 'LEFT':
            s = LEFT
        if style == 'RIGHT':
            s = RIGHT
        if style == 'BLOCK':
            s = BLOCK
        self.cursor.ParaAdjust = s
        print(self.cursor.ParaAdjust)
        return 1


def run_xmlrpc_server():
    port = 1210
    word_proxy = WordProxy()
    server = SimpleXMLRPCServer(("localhost", port))
    server.register_instance(word_proxy)
    server.serve_forever()


if __name__ == '__main__':
    print("running server...")
    run_xmlrpc_server()
    
    print("server exited")
