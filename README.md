# VBscriptInternals

Author: [Boris Larin](https://twitter.com/oct0xor)

This repository contains
scripts for disassembling VBScript p-code in the memory to aid in exploits
analysis.

https://securelist.com/delving-deep-into-vbscript-analysis-of-cve-2018-8174-exploitation/86333/

## Contents

`kl_vbs_disasm_ida.py` - Script for IDA Pro 

`kl_vbs_disasm_windbg.py` - Script for WinDbg with PyKD extension

## Usage

Set breakpoint at
function `vbscript!CScriptRuntime::RunNoEH` and use appropriate script after breakpoint is hit. 
