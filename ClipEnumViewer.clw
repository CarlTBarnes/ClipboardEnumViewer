!Example program to Enumerate and Get Data from Clipboard by Carl Barnes. Written circa 2000
!Released under MIT License

    PROGRAM
    INCLUDE('KEYCODES.CLW') 
!Region Equates    
CF_FIRST            EQUATE(0)
CF_TEXT             EQUATE(1)

CF_BITMAP           EQUATE(2)
CF_METAFILEPICT     EQUATE(3)
CF_SYLK             EQUATE(4)
CF_DIF              EQUATE(5)
CF_OEMTEXT          EQUATE(7)
CF_DIB              EQUATE(8)
CF_PALETTE          EQUATE(9)
CF_PENDATA          EQUATE(10)
CF_RIFF             EQUATE(11)
CF_WAVE             EQUATE(12)
CF_UNICODETEXT      EQUATE(13)
CF_ENHMETAFILE      EQUATE(14)
CF_HDROP            EQUATE(15)
CF_LOCALE           EQUATE(16)
CF_MAX              EQUATE(17)

CF_PRIVATEFIRST     EQUATE(0200h)
CF_PRIVATELAST      EQUATE(02FFh)

CF_GDIOBJFIRST      EQUATE(0300h)
CF_GDIOBJLAST       EQUATE(03FFh)
CF_OWNERDISPLAY     EQUATE(0080h)
CF_DSPTEXT          EQUATE(0081h)
CF_DSPBITMAP        EQUATE(0082h)
CF_DSPENHMETAFILE   EQUATE(008Eh)

!HANDLE                  EQUATE(UNSIGNED)
!HWND                    EQUATE(HANDLE)
!HGLOBAL                 EQUATE(HANDLE)
!BOOL                    EQUATE(SIGNED)
!LPCSTR                  EQUATE(CSTRING)  
!LPSTR                   EQUATE(CSTRING)    !Usage:Pass the Label of the LPSTR
!WORD                    EQUATE(SIGNED)
!DWORD                   EQUATE(ULONG)
!EndRegion Equates

    MAP 
ClipEnumViewer  PROCEDURE()    
GetCBFormatName PROCEDURE(LONG pFormat_CB_No),STRING  !Return Pretty name for API CB_Define
  
       MODULE('Windows.DLL')
           GetClipboardFormatName(UNSIGNED,*CSTRING,SIGNED),SIGNED,RAW,PASCAL,NAME('GetClipboardFormatNameA'),DLL(1)
           EnumClipboardFormats(UNSIGNED),UNSIGNED,PASCAL,DLL(1)
           OpenClipboard(LONG HWND),BOOL,PASCAL,DLL(1)
           CloseClipboard(),BOOL,PASCAL,DLL(1)
           GetLastError(),LONG,PASCAL ,DLL(1)
           
!Other API CLipboard functions
!           GetClipboardData(UNSIGNED),UNSIGNED,PASCAL,DLL(1)
!           GetClipboardOwner(),LONG,PASCAL,DLL(1)
!           GetOpenClipboardWindow(),LONG,PASCAL,DLL(1)
!           GlobalLock(LONG HGLOBAL),LONG,PASCAL,DLL(1)
!           GlobalUnlock(LONG HGLOBAL),BOOL,PASCAL,DLL(1)
!           GlobalSize(LONG HGLOBAL),DWORD,PASCAL,DLL(1)     
!           GlobalHandle(LONG),LONG,PASCAL,RAW,DLL(1)  
       END
    END !MAP

    CODE
    SYSTEM{PROP:PropVScroll}=1
    SYSTEM{PROP:MsgModeDefault}=MSGMODE:CANCOPY
    ClipEnumViewer()
    RETURN 

!============================================
ClipEnumViewer  PROCEDURE()    
HCBOwner    LONG
HCBOpen     LONG

CBFormat    UNSIGNED
CustomCB    UNSIGNED
CustomName  CSTRING(999)

!ShowText    STRING(2000)
GetCbNumber     LONG    
MxSize          EQUATE(32000) 
ShowCbContent   STRING(MxSize)    
ShowBin         BYTE(1) 
Show0           BYTE(0)  

EnumFormatQ     QUEUE,PRE(EnumQ)
Number              LONG 
Name                STRING(255)
Sample              STRING(255) 
            END              
EnumsAsText STRING(4000)

ShowSampleData  BOOL             
HScrollTxt      BYTE(1)
            
Window WINDOW('Clipboard Enum and View'),AT(,,306,265),GRAY,SYSTEM,ICON(ICON:Paste),FONT('Segoe UI',10), |
            RESIZE
        LIST,AT(3,3,,90),FULL,USE(?List:EnumFormatQ),VSCROLL,FROM(EnumFormatQ),FORMAT('36R(2)|FM~For' & |
                'mat#~L(2)@n_7@112L(2)|M~Format Name~@s255@Q''From GetClipboardFormatName() or CB_De' & |
                'fine Name''20L(2)~Sample  -  Double Click to Select~@s255@'),ALRT(DeleteKey)
        BUTTON('&Refresh'),AT(3,97,48,14),USE(?RefreshBtn),TIP('EnumClipboardFormats')
        CHECK('Sample Data'),AT(62,99),USE(ShowSampleData),TIP('Show Sample Data Column. May crash?')
        BUTTON('&Copy List **'),AT(152,97,,14),USE(?CopyListBtn),SKIP,TIP('Copy List of Formats to C' & |
                'lipboard<13,10><13,10>** FYI You Lose Clipboard')
        PROMPT('Open Word or Excel and Copy to the Clipboard to see a lot of formats.'),AT(221,97,82,23), |
                USE(?FYI),FONT(,9)
        BUTTON('&GetFormat #'),AT(3,120,48),USE(?GetFormatBtn),TIP('Calls Clarion Clipboard( # ) for' & |
                ' Number entered.<13,10>Or double-click on list of formats.')
        ENTRY(@n7),AT(62,122,36),USE(GetCbNumber)
        TEXT,AT(4,139),FULL,USE(ShowCbContent),HVSCROLL
        CHECK('<<&Binary>'),AT(106,125,40),USE(ShowBin),SKIP,TIP('Show 1-31 and 127-255 as <<#>.')
        CHECK('<<&x0>'),AT(155,125),USE(Show0),SKIP,TIP('Show hex 0 as <<x0> else shows blank.')
        CHECK('HScroll'),AT(194,125),USE(HScrollTxt),SKIP
    END
    CODE
    OPEN(Window)
    0{PROP:MinWidth}=0{PROP:Width} 
    0{PROP:MinHeight}=0{PROP:Height} - ?ShowCbContent{PROP:Height} * .60
    ACCEPT
        CASE EVENT()
        OF EVENT:OpenWindow ; POST(EVENT:Accepted,?RefreshBtn) 
        END 
        CASE ACCEPTED()
        OF ?RefreshBtn      ; DO LoadEnumFormatQRtn
        OF ?GetFormatBtn    ; DO GetFormatRtn  
        OF ?GetCbNumber     ; DO GetFormatRtn
        OF ?HScrollTxt      ; ?ShowCbContent{PROP:HScroll}=HScrollTxt  
        OF ?CopyListBtn     ; SetClipboard(EnumsAsText) 
        END  
        CASE FIELD()
        OF ?List:EnumFormatQ
           GET(EnumFormatQ,CHOICE(?List:EnumFormatQ))
           CASE EVENT()
           OF EVENT:AlertKey
              IF KEYCODE()=MouseLeft2 THEN DELETE(EnumFormatQ).
           OF EVENT:NewSelection 
              IF KEYCODE()=MouseLeft2 THEN
                 GetCbNumber = EnumQ:Number
                 DO GetFormatRtn
              END 
           END
        END 
    END               
    RETURN    
!---------------------
GetFormatRtn ROUTINE
    DATA
BinTxt  STRING(MxSize-7),AUTO   
BX      LONG  
OutX    LONG  
MaxX    LONG  
AX      LONG 
AByte   PSTRING(6),AUTO    
    CODE   
    IF GetCbNumber=0 THEN
        ShowCbContent = 'Meet me half way and fill in the Format Number you wish to display.'
        DISPLAY
        SELECT(?GetCbNumber)
        EXIT
    END 
    IF ~ShowBin THEN 
        ShowCbContent = CLIPBOARD(GetCbNumber) 
        DISPLAY
        EXIT
    END
    BinTxt = CLIPBOARD(GetCbNumber) 
    IF ~BinTxt THEN                         !Blank data on clipboard?
        EnumQ:Number = GetCbNumber
        GET(EnumFormatQ,EnumQ:Number)
       IF ERRORCODE() THEN
          ShowCbContent='Format #'& GetCbNumber  &' is NOT in the list <13,10>Blank was returned by CLIPBOARD().'
          DISPLAY 
          EXIT 
       END 
    ELSIF ~ERRORCODE() THEN 
    END 
    ShowCbContent=''
    MaxX = SIZE(ShowCbContent)
    Loop BX=1 TO LEN(CLIP(BinTxt)) 
         IF OutX + 1 > MaxX THEN BREAK. 
         CASE VAL(BinTxt[BX])
         OF 32 TO 126   ; OutX += 1 ; ShowCbContent[OutX]=BinTxt[BX] ; CYCLE 
         OF 0           ; IF ~VAL(BinTxt[BX]) AND ~Show0 ; OutX += 1 ; CYCLE.    !blank zeros
         END
         AByte='<' & val(BinTxt[BX]) &'>'
         LOOP AX=1 TO LEN(AByte)
              IF OutX + 1 > MaxX THEN BREAK.
              OutX += 1 
              ShowCbContent[OutX]=AByte[AX]  
         END 
    WHILE OutX < MxSize

    EnumQ:Number = GetCbNumber           !Sync List to number entered
    GET(EnumFormatQ,EnumQ:Number)
    IF ~ERRORCODE() THEN
       EnumQ:Sample=ShowCbContent        !and show data in sample column
       PUT(EnumFormatQ)
       ?List:EnumFormatQ{PROP:Selected}=POINTER(EnumFormatQ) 
    END 
    DISPLAY
    EXIT

!-------------------------------------------------------------------
LoadEnumFormatQRtn ROUTINE
    DATA
LnClp   LONG
    CODE
    FREE(EnumFormatQ)
    EnumsAsText = 'Format#<9>Format Name<13,10>'
    IF ~OpenClipboard(0{PROP:Handle}) THEN 
       MESSAGE('OpenClipboard() failed||Last Error=' & GetLastError(),'LoadClipRtn ROUTINE')
       DISPLAY
       EXIT
    END

    CBFormat = CF_First
    LOOP
        CBFormat = EnumClipboardFormats(CBFormat)
        IF ~CBFormat THEN BREAK.
       
        CLEAR(EnumFormatQ) 
        EnumQ:Number = CBFormat 

        LnClp = GetClipboardFormatName(CBFormat, CustomName, SIZE(CustomName)-1 )
        IF ~LnClp THEN
           CustomName=GetCBFormatName(CBFormat)
        END
        EnumQ:Name = CustomName
        ADD(EnumFormatQ,EnumQ:Number) 
        EnumsAsText = CLIP(EnumsAsText) &  CBFormat &'<9>'&  CLIP(CustomName) & '<13,10>'
    END

    IF ~CloseClipboard() THEN                                         !0=failed
       Message('CloseClipboard() failed||Last Error=' & GetLastError() ,'LoadClipRtn ROUTINE')
    END
    DISPLAY
    IF ShowSampleData THEN DO ShowSampleDataRtn.
    ?List:EnumFormatQ{PROP:Selected}=1
    IF GetCbNumber    THEN 
        DO SyncGetCbNumberRtn
    END 
    DISPLAY
    EXIT

SyncGetCbNumberRtn ROUTINE       
    EnumQ:Number = GetCbNumber
    GET(EnumFormatQ,EnumQ:Number)
    IF ERRORCODE() THEN
       GetCbNumber=0
       CLEAR(ShowCbContent)
       EXIT
    END 
    ?List:EnumFormatQ{PROP:Selected}=POINTER(EnumFormatQ) 
    EXIT

ShowSampleDataRtn ROUTINE
    DATA
QX              LONG   
SaveCbNumber    LONG   
    CODE
    SaveCbNumber  = GetCbNumber
    LOOP QX=1 TO RECORDS(EnumFormatQ)
        GET(EnumFormatQ,QX)
        ?List:EnumFormatQ{PROP:Selected}=QX
        GetCbNumber = EnumQ:Number
        DO GetFormatRtn           
        EnumQ:Sample = ShowCbContent
        PUT(EnumFormatQ) 
        DISPLAY
    END
    IF SaveCbNumber THEN
       GetCbNumber = SaveCbNumber
       DO GetFormatRtn     
    END 
    EXIT

!================================================================
GetCBFormatName PROCEDURE(LONG pFormat_CB_No)!,STRING
CbName  STRING(32)
    CODE
    CASE pFormat_CB_No
    OF CF_TEXT             ; CbName = 'CF_TEXT '
    OF CF_BITMAP           ; CbName = 'CF_BITMAP '
    OF CF_METAFILEPICT     ; CbName = 'CF_METAFILEPICT '
    OF CF_SYLK             ; CbName = 'CF_SYLK '
    OF CF_DIF              ; CbName = 'CF_DIF '
    OF CF_OEMTEXT          ; CbName = 'CF_OEMTEXT '
    OF CF_DIB              ; CbName = 'CF_DIB '
    OF CF_PALETTE          ; CbName = 'CF_PALETTE '
    OF CF_PENDATA          ; CbName = 'CF_PENDATA '
    OF CF_RIFF             ; CbName = 'CF_RIFF '
    OF CF_WAVE             ; CbName = 'CF_WAVE '
    OF CF_UNICODETEXT      ; CbName = 'CF_UNICODETEXT '
    OF CF_ENHMETAFILE      ; CbName = 'CF_ENHMETAFILE '
    OF CF_HDROP            ; CbName = 'CF_HDROP '
    OF CF_LOCALE           ; CbName = 'CF_LOCALE '
    OF CF_MAX              ; CbName = 'CF_MAX '

    OF CF_PRIVATEFIRST TO CF_PRIVATELAST      ; CbName = 'CF_PRIVATE ' !use GetClipboardFormatName
    OF CF_GDIOBJFIRST  TO CF_GDIOBJLAST       ; CbName = 'CF_GDIOB'

    OF CF_OWNERDISPLAY     ; CbName = 'CF_OWNERDISPLAY  '
    OF CF_DSPTEXT          ; CbName = 'CF_DSPTEXT '
    OF CF_DSPBITMAP        ; CbName = 'CF_DSPBITMAP '
    OF CF_DSPENHMETAFILE   ; CbName = 'CF_DSPENHMETAFILE '

    ELSE
        CbName = '? #' & pFormat_CB_No
    END    
    
    RETURN CLIP(CbName) 