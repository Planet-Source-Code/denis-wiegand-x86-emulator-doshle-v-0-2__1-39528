Attribute VB_Name = "Interrups"
'x86 HLE Emultor @ 2002 Project by Denis Wiegand
'Short Storry: After i stoped the devolopent on my "VB ModPlayer"
'i started with this new and fresh idea. It`s great to Emulate old DOS
'games on a high-End Multimedia P4 Engine. It IS NOT easy to dev. this.
'What can be Emulated:
'BIOS->Text Interrupts
'BASIC CPU-operations (most w/o R/M byte) (ALL R/M byte`s are currently only HLE emulated) see below.
'This is a great caos of code with and not used segments of text like CPU.bas.
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

'All System Interrups
Dim tmp As Integer
Global screenX As Long
Global screenY As Long
'BIOS

Function x86_int(ByVal num As Byte)
Select Case num

    Case 33         'Text operations
        Select Case ah
        
        Case 9      'type of operiaton "print, read,..."
            While Mem((CS + (dx - 256)) + tmp) <> Asc("$") 'Seek to "$"
                frmMain.PrintDOS Mem((CS + (dx - 256)) + tmp)   'Simple Textoutput
                tmp = tmp + 1
            Wend
            tmp = 0
        Case 2
            frmMain.PrintDOS dl
        Case 76
            brun = False
        End Select
    Case 16
        Select Case ah
        Case 9
            For I = 0 To cx
                frmMain.PrintDOS al
            Next I
        End Select
         Select Case ah
        Case 14
            For I = 0 To cx
                frmMain.PrintDOS al
            Next I
        End Select
    End Select
    
End Function
