Attribute VB_Name = "Emulator"
'x86 HLE Emultor @ 2002 Project by Denis Wiegand
'Short Storry: After i stoped the devolopent on my "VB ModPlayer"
'i started with this new and fresh idea. It`s great to Emulate old DOS
'games on a high-End Multimedia P4 Engine. It IS NOT easy to dev. this.
'What can be Emulated:
'BIOS->Text Interrupts
'BASIC CPU-operations (most w/o R/M byte) (ALL R/M byte`s are currently only HLE emulated) see below.
'All Emulation is done in "Emulate()" and the code is loaded into "Mem" with a space of 1179648 byte.
'Enough ...

'Short description of H.L.E
'HLE means "High Level Emulator" that emulates
'blocks of code with one function. Ok a simple
'example: You Call The function >MOV EBP,(edi+0xffff)<
'ok in the compiled binary are the function`s code for "MOV"
'now the next following is an (Register/Memory) byte that select`s
'a register or an Memory location from a Table. The HLE dont need
'an table he rips the full code with R/M byte as one opcode!! and emulates them
'See UltraHLE (N64) emualtor as example. (I hope that descripes the HLE function)

'Sorry for my *stupid* english.


Global AsmText() As Byte  'Into Memory
Global brun As Boolean  'To Stop emulation ;)
Dim lo As Long
Dim hi As Long
Global ips As Long

Dim ZF As Boolean

Function Emulate()
IP = 0
SI = 0
BP = 0
DI = 0
VES = 0
al = 0
ah = 0
ax = 0
bl = 0
bh = 0
bx = 0
cl = 0
ch = 0
cx = 0
dl = 0
dh = 0
dx = 0
SP = (SS + 65536)
VES = ES
screenX = 0
screenY = 0
frmMain.Picture1.Cls
brun = True
frmMain.StatusBar1.Panels(1).Text = "**Running**"
Do While brun
DoEvents
here:

    'For 0 opcodes goes here
    If Len(Hex(Mem(CS + IP))) = 1 Then
        Select Case Mid(Hex(Mem(CS + IP)), 1, 2)
        Case "4"
        'ADD AL, Vb
        IP = IP + 1
        al = (al + (Mem(CS + IP)))
        IP = IP + 1
        Case "5"
        'ADD AX, Vv
        IP = IP + 1
        hi = Mem(CS + IP)
        IP = IP + 1
        lo = Mem(CS + IP)
        ax = (ax + ((loo * 256) + hii))
        IP = IP + 1
        Case "C"
        'OR AL, Vb
        IP = IP + 1
        al = (al Or Mem(CS + IP))
        Case "D"
        IP = IP + 1
        hi = Mem(CS + IP)
        IP = IP + 1
        lo = Mem(CS + IP)
        ax = (ax Or ((loo * 256) + hii))
        IP = IP + 1
        End Select
    End If

    Select Case Mid(Hex(Mem(CS + IP)), 1, 1)    'Opcodes beginning
    
    Case "2"    '2
        Select Case Mid(Hex(Mem(CS + IP)), 2, 1)
        Case "4"
        'AND AL, Vb
        IP = IP + 1
        al = (al And Mem(CS + IP))
        IP = IP + 1
        Case "5"
        'AND AX, Vv
        IP = IP + 1
        hi = Mem(CS + IP)
        IP = IP + 1
        lo = Mem(CS + IP)
        ax = (ax And ((lo * 256) + hi))
        IP = IP + 1
        'Extendet R/M HLE-emulation>
        Case "E"
           Select Case Mid(Hex(Mem(CS + IP + 1)), 1, 2)
           Case "8A"
                Select Case Mid(Hex(Mem(CS + IP + 2)), 1, 2)
                    Case "14"
                        dl = Mem(CS + SI)
                        IP = IP + 3
                    Case "4"
                        al = Mem(CS + SI)
                        IP = IP + 3
                End Select
            Case "89"
                Select Case Mid(Hex(Mem(CS + IP + 2)), 1, 2)
                    Case "36"
                    IP = IP + 3
                    hi = Mem(CS + IP)
                    IP = IP + 1
                    lo = Mem(CS + IP)
                    Mem((CS + ((lo * 256) + hi))) = SI
                    IP = IP + 1
                 End Select
            Case "8B"
                Select Case Mid(Hex(Mem(CS + IP + 2)), 1, 2)
                    Case "36"
                    IP = IP + 3
                    hi = Mem(CS + IP)
                    IP = IP + 1
                    lo = Mem(CS + IP)
                    SI = Mem((CS + ((lo * 256) + hi)))
                    IP = IP + 1
            End Select
        End Select
        End Select
        
        
   Case "3"     '3
        Select Case Mid(Hex(Mem(CS + IP)), 2, 1)
        Case "4"
        'XOR AL, Vb
        IP = IP + 1
        al = (al Xor Mem(CS + IP))
        IP = IP + 1
        Case "5"
        'XOR AX, Vv
        IP = IP + 1
        hi = Mem(CS + IP)
        IP = IP + 1
        lo = Mem(CS + IP)
        ax = (ax Xor ((lo * 256) + hi))
        IP = IP + 1
        Case "C"
                ZF = (Mem(CS + IP + 1) = al)
                IP = IP + 2
        Case "B"
                Select Case Mid(Hex(Mem(CS + IP + 1)), 1, 2)
                Case "16"
                IP = IP + 1
                hi = Mem(CS + IP)
                IP = IP + 1
                lo = Mem(CS + IP)
                ZF = (Mem((lo * 256) + hi) = dx)
                End Select
        End Select
        
    Case "4"    '4
        Select Case Mid(Hex(Mem(CS + IP)), 2, 1)
        
        Case "0"
        'INC AX
        ax = ax + 1
        IP = IP + 1
        Case "1"
        'INC CX
        cx = cx + 1
        IP = IP + 1
        Case "2"
        'INC DX
        dx = dx + 1
        IP = IP + 1
        Case "3"
        'INC BX
        bx = bx + 1
        IP = IP + 1
        Case "4"
        'INC SP
        SP = SP + 1
        IP = IP + 1
        Case "5"
        'INC BP
        BP = BP + 1
        IP = IP + 1
        Case "6"
        'INC SI
        SI = SI + 1
        IP = IP + 1
        Case "7"
        'INC DI
        DI = DI + 1
        IP = IP + 1
        
        Case "8"
        'DEC AX
        ax = ax - 1
        IP = IP + 1
        Case "9"
        'DEC CX
        cx = cx - 1
        IP = IP + 1
        Case "A"
        'DEC DX
        dx = dx - 1
        IP = IP + 1
        Case "B"
        'DEC BX
        bx = bx - 1
        IP = IP + 1
        Case "C"
        'DEC SP
        SP = SP - 1
        IP = IP + 1
        Case "D"
        'DEC BP
        BP = BP - 1
        IP = IP + 1
        Case "E"
        'DEC SI
        SI = SI - 1
        IP = IP + 1
        Case "F"
        'DEC DI
        DI = DI - 1
        IP = IP + 1
        End Select
    
    Case "5"    '5
    'PUSH and POP
        Select Case Mid(Hex(Mem(CS + IP)), 2, 1)
        Case "0"
            SP = SP - 2
            Mem(SS + SP + 1) = (ax Mod 256)
            Mem(SS + SP + 2) = (ax \ 256)
            IP = IP + 1
        Case "1"
            SP = SP - 2
            Mem(SS + SP + 1) = (cx Mod 256)
            Mem(SS + SP + 2) = (cx \ 256)
            IP = IP + 1
        Case "2"
            SP = SP - 2
            Mem(SS + SP + 1) = (dx Mod 256)
            Mem(SS + SP + 2) = (dx \ 256)
            IP = IP + 1
        Case "3"
            SP = SP - 2
            Mem(SS + SP + 1) = (bx Mod 256)
            Mem(SS + SP + 2) = (bx \ 256)
            IP = IP + 1
        Case "4"
            SP = SP - 2
            Mem(SS + SP + 1) = (SP Mod 256)
            Mem(SS + SP + 2) = (SP \ 256)
            IP = IP + 1
        Case "5"
            SP = SP - 2
            Mem(SS + SP + 1) = (BP Mod 256)
            Mem(SS + SP + 2) = (BP \ 256)
            IP = IP + 1
        Case "6"
            SP = SP - 2
            Mem(SS + SP + 1) = (SI Mod 256)
            Mem(SS + SP + 2) = (SI \ 256)
            IP = IP + 1
        Case "7"
            SP = SP - 2
            Mem(SS + SP + 1) = (DI Mod 256)
            Mem(SS + SP + 2) = (DI \ 256)
            IP = IP + 1
        Case "8"
            IP = IP + 1
            hi = Mem(SS + SP + 1)
            lo = Mem(SS + SP + 2)
            ax = ((lo * 256) + hi)
            SP = SP + 2
        Case "9"
            IP = IP + 1
            hi = Mem(SS + SP + 1)
            lo = Mem(SS + SP + 2)
            cx = ((lo * 256) + hi)
            SP = SP + 2
        Case "A"
            IP = IP + 1
            hi = Mem(SS + SP + 1)
            lo = Mem(SS + SP + 2)
            dx = ((lo * 256) + hi)
            SP = SP + 2
        Case "B"
            IP = IP + 1
            hi = Mem(SS + SP + 1)
            lo = Mem(SS + SP + 2)
            bx = ((lo * 256) + hi)
            SP = SP + 2
        Case "C"
            IP = IP + 1
            hi = Mem(SS + SP + 1)
            lo = Mem(SS + SP + 2)
            SP = ((lo * 256) + hi)
            SP = SP + 2
        Case "D"
            IP = IP + 1
            hi = Mem(SS + SP + 1)
            lo = Mem(SS + SP + 2)
            BP = ((lo * 256) + hi)
            SP = SP + 2
        Case "E"
            IP = IP + 1
            hi = Mem(SS + SP + 1)
            lo = Mem(SS + SP + 2)
            SI = ((lo * 256) + hi)
            SP = SP + 2
        Case "F"
            IP = IP + 1
            hi = Mem(SS + SP + 1)
            lo = Mem(SS + SP + 2)
            DI = ((lo * 256) + hi)
            SP = SP + 2
        End Select

    Case "9"    '9
    'XCHG Register,AX
        Select Case Mid(Hex(Mem(CS + IP)), 2, 1)
        Case "0"
        'NO OPERATION "NOP"
        Case "1"
        z = cx
        cx = ax
        ax = z
        IP = IP + 1
        Case "2"
        z = cx
        dx = ax
        ax = z
        IP = IP + 1
        Case "3"
        z = bx
        bx = ax
        ax = z
        IP = IP + 1
        Case "4"
        z = SP
        SP = ax
        ax = z
        IP = IP + 1
        Case "5"
        z = BP
        BP = ax
        ax = z
        IP = IP + 1
        Case "6"
        z = SI
        SI = ax
        ax = SI
        IP = IP + 1
        Case "7"
        z = DI
        DI = ax
        ax = z
        IP = IP + 1
        End Select

    Case "A"    '10
        Select Case Mid(Hex(Mem(CS + IP)), 2, 1)
        Case "0"
            'Mov al, offset
            IP = IP + 1
            lo = Mem(CS + IP)
            IP = IP + 1
            hi = Mem(CS + IP)
            al = Mem(CS + ((lo * 256) + hi))
            IP = IP + 1
        Case "1"
            'Mov ax, offset
            IP = IP + 1
            lo = Mem(CS + IP)
            IP = IP + 1
            hi = Mem(CS + IP)
            loo = Mem(CS + ((lo * 256) + hi) + 1)
            hii = Mem(CS + ((lo * 256) + hi))
            ax = ((loo * 256) + hii)
            IP = IP + 1
        Case "2"
            'Mov offset, al
            IP = IP + 1
            lo = Mem(CS + IP)
            IP = IP + 1
            hi = Mem(CS + IP)
            Mem(CS + ((lo * 256) + hi)) = al
            IP = IP + 1
        Case "3"
            'Mov offset, ax
            IP = IP + 1
            lo = Mem(CS + IP)
            IP = IP + 1
            hi = Mem(CS + IP)
            
            Mem((CS + ((lo * 256) + hi))) = (ax Mod 256)
            Mem((CS + ((lo * 256) + hi)) + 1) = (ax \ 256)
            IP = IP + 1
        End Select


        


    Case "B"    '11
        Select Case Mid(Hex(Mem(CS + IP)), 2, 1)
        Case "0"
            'It`s for AL
            IP = IP + 1
            al = (Mem(CS + IP))
            IP = IP + 1
        Case "1"
            IP = IP + 1
            cl = (Mem(CS + IP))
            IP = IP + 1
        Case "2"
            IP = IP + 1
            dl = (Mem(CS + IP))
            IP = IP + 1
        Case "3"
            'It`s for BL
            IP = IP + 1
            bl = (Mem(CS + IP))
            IP = IP + 1
        Case "4"
            'It`s for AH
            IP = IP + 1
            ah = (Mem(CS + IP))
            IP = IP + 1
        Case "5"
            'It`s for CH
            IP = IP + 1
            ch = (Mem(CS + IP))
            IP = IP + 1
        Case "6"
            'It`s for DH
            IP = IP + 1
            dh = (Mem(CS + IP))
            IP = IP + 1
        Case "7"
            'It`s for BH
            IP = IP + 1
            bh = (Mem(CS + IP))
            IP = IP + 1
        Case "8"
            'It`s for AX
            IP = IP + 1
            hi = Mem(CS + IP)
            IP = IP + 1
            lo = Mem(CS + IP)
            ax = ((lo * 256) + hi)
            IP = IP + 1
        Case "9"
            'It`s for CX
            IP = IP + 1
            hi = Mem(CS + IP)
            IP = IP + 1
            lo = Mem(CS + IP)
            cx = ((lo * 256) + hi)
            IP = IP + 1
        Case "A"
            'It`s for DX
            IP = IP + 1
            hi = Mem(CS + IP)
            IP = IP + 1
            lo = Mem(CS + IP)
            dx = ((lo * 256) + hi)
            IP = IP + 1
        Case "B"
            'It`s for BX
            IP = IP + 1
            hi = Mem(CS + IP)
            IP = IP + 1
            lo = Mem(CS + IP)
            bx = ((lo * 255) + hi)
            IP = IP + 1
        Case "C"
            'It`s for SP
            IP = IP + 1
            hi = Mem(CS + IP)
            IP = IP + 1
            lo = Mem(CS + IP)
            SP = ((lo * 255) + hi)
            IP = IP + 1
        Case "D"
            'It`s for BP
            IP = IP + 1
            hi = Mem(CS + IP)
            IP = IP + 1
            lo = Mem(CS + IP)
            BP = ((lo * 255) + hi)
            IP = IP + 1
        Case "E"
            'It`s for SI
            IP = IP + 1
            hi = Mem(CS + IP)
            IP = IP + 1
            lo = Mem(CS + IP)
            SI = ((lo * 255) + hi)
            IP = IP + 1
        Case "F"
            IP = IP + 1
            hi = Mem(CS + IP)
            IP = IP + 1
            lo = Mem(CS + IP)
            DI = ((lo * 256) + hi)
            IP = IP + 1
        End Select

    Case "C"
        Select Case Mid(Hex(Mem(CS + IP)), 2, 1)
        Case "D"
            'It`s a Interrupt
            IP = IP + 1
            x86_int (Mem(CS + IP))
            IP = IP + 1
        Case "3"
            'Ret "Return"
            If SP = 65536 Then Exit Do 'We return to DOS ;)
            IP = IP + 1
            hi = Mem(SS + SP + 1)
            lo = Mem(SS + SP + 2)
            SP = SP + 2
            IP = ((lo * 256) + hi)
        End Select
        
    Case "E"
         Select Case Mid(Hex(Mem(CS + IP)), 2, 1)
         Case "8"
            'Call "FUNCTION"
            IP = IP + 1
            hi = Mem(CS + IP)
            IP = IP + 1
            lo = Mem(CS + IP)
            IP = IP + 1
            SP = SP - 2
            Mem(SS + SP + 1) = (IP Mod 256)
            Mem(SS + SP + 2) = (IP \ 256)
            IP = ((lo * 256) + hi) + IP
        Case "9"
            IP = IP + 1
            hi = Mem(CS + IP)
            IP = IP + 1
            lo = Mem(CS + IP)
            IP = ((((lo * 256) + hi) - 65535) + IP)
            
        Case "B"
            'Jump short
            If Int(Mem(CS + IP + 1) / 16) = 15 Then 'We Going backwards
                IP = ((IP - (Mem(CS + IP + 1) \ 16)) + ((Mem(CS + IP + 1) Mod 16) + 1))
                
            Else
            IP = (IP + (Mem(CS + IP + 1) + 2))  'We Going forwards
            End If
        Case "2"
            If cx > 0 Then
              If Int(Mem(CS + IP + 1) / 16) = 15 Then
                    IP = ((IP - (Mem(CS + IP + 1) \ 16)) + ((Mem(CS + IP + 1) Mod 16) + 1))
              Else
                    IP = (Mem(CS + IP + 1) + 2)
              End If
              cx = cx - 1
            Else
              IP = IP + 2
            End If
        End Select

    'Case "2"

        'End Select
        
    Case "8"
        Select Case Mid(Hex(Mem(CS + IP)), 2, 1)
        Case "0"
            Select Case Mid(Hex(Mem(CS + IP + 1)), 1, 2)
            Case "FA"
                ZF = (Mem(CS + IP + 2) = dl)
                IP = IP + 3
            Case "3E"
                IP = IP + 2
                hi = Mem(CS + IP)
                IP = IP + 1
                lo = Mem(CS + IP)
                IP = IP + 1
                ZF = Mem(((lo * 256) + hi)) = Mem(CS + IP)
                IP = IP + 1
            End Select
        Case "A"
            Select Case Mid(Hex(Mem(CS + IP + 1)), 1, 2)
            Case "60"
                IP = IP + 2
                hi = Mem(CS + IP)
                IP = IP + 1
                lo = Mem(CS + IP)
                al = Mem((lo * 256) + hi)
            End Select
        Case "B"
            Select Case Mid(Hex(Mem(CS + IP + 1)), 1, 2)
            Case "16"
                IP = IP + 2
                lo = Mem(CS + IP)
                IP = IP + 1
                hi = Mem(CS + IP)
                dx = ((lo * 256) + hi)
                IP = IP + 1
            Case "85"
                IP = IP + 2
                hi = Mem(CS + IP)
                IP = IP + 1
                lo = Mem(CS + IP)
                ax = (DI + ((lo * 256) + hi))
                IP = IP + 1
            End Select
        Case "9"
            Select Case Mid(Hex(Mem(CS + IP + 1)), 1, 2)
            Case "85"
                IP = IP + 2
                hi = Mem(CS + IP)
                IP = IP + 1
                lo = Mem(CS + IP)
                Mem((DI + ((lo * 256) + hi))) = (ax \ 256)
                Mem((DI + ((lo * 256) + hi)) + 1) = (ax Mod 256)
                IP = IP + 1
            Case "16"
                IP = IP + 2
                lo = Mem(CS + IP)
                IP = IP + 1
                hi = Mem(CS + IP)
                Mem((DI + ((hi * 256) + lo))) = (dx \ 256)
                Mem((DI + ((hi * 256) + lo)) + 1) = (dx Mod 256)
                IP = IP + 1
            End Select
        Case "8"
            Select Case Mid(Hex(Mem(CS + IP + 1)), 1, 2)
            Case "26"
                IP = IP + 2
                lo = Mem(CS + IP)
                IP = IP + 1
                hi = Mem(CS + IP)
                Mem((DI + ((hi * 256) + lo))) = ah
                IP = IP + 1
            End Select
        Case "3"
            Select Case Mid(Hex(Mem(CS + IP + 1)), 1, 2)
            Case "EF"
                IP = IP + 2
                lo = Mem(CS + IP)
                IP = IP + 1
                hi = Mem(CS + IP)
                DI = DI - Mem(CS + IP)
                IP = IP + 1
            Case "C2"
                IP = IP + 2
                dx = dx + Mem(CS + IP)
                IP = IP + 1
            End Select
        End Select
    

    Case "7"
        Select Case Mid(Hex(Mem(CS + IP)), 2, 1)
        Case "4"
            If ZF Then
                IP = (IP + Mem(CS + IP + 1)) + 2
            Else
                IP = IP + 2
            End If
        Case "2"
            'Do Nothing
            IP = (IP + Mem(CS + IP + 1)) + 2
            IP = IP + 2
         End Select
    Case Else
        'An unknown opcode goes here
        MsgBox "Unknown Opcode at: " & Hex(CS + IP) & vbCrLf & "Value: " & Hex(Mem(CS + IP)), vbCritical, "Error"
        Exit Do
    End Select
Loop
frmMain.StatusBar1.Panels(1).Text = "**Stop**"
End Function

