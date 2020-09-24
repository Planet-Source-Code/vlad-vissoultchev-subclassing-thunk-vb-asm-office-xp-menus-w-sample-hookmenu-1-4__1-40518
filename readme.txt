first register this typelib

	1_Tlb\SubclassingSink.tlb
	
then open some of the sample projects

	4_Samples\1_Simple\SubclassingThunk.vbp
	4_Samples\2_HookMenu\Sample\Group1.vbp
	4_Samples\3_OutlookBar\OutlookBar.vbp 

if you need the subclassing thunk in a project of yours just add this file

	3_Thunks\cSubclassingThunk.cls 
	
add a reference to this typelib
	
	"Subclassing/Hooking sink interfaces 1.0" (SubclassingSink.tlb)
	
and in a form or class in your project implement this interface

	ISubclassingSink 
	
if you need the hooking thunk in a project of yours just add this file

	3_Thunks\cHookingThunk.cls 

you can take a look at the underlying assembly in

	2_Asm\WndProc2.asm 
	2_Asm\HookProc.asm 

you will probably need the assembler/linker/editor and tools from http://www.masm32.com

enjoy,
</wqw>
8 Nov 2002