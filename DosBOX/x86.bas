Attribute VB_Name = "CPU"
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

'AX     'Accumulator
Global al As Long
Global ah As Long
Global ax As Long

'BX     'Baseregister
Global bl As Long
Global bh As Long
Global bx As Long

'CX     'CountRegister
Global cl As Long
Global ch As Long
Global cx As Long

'DX     'Dataregister
Global dl As Long
Global dh As Long
Global dx As Long

Global IP As Long 'for CodeSegment > CS:IP

Function x86_Interrupt(num As Integer)

End Function

Function x86_mov(ByRef a As Long, ByVal b As Long)
a = b   'Mov is the simpelst function
End Function

