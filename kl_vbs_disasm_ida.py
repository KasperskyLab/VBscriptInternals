# VBScript VM disassembler 
#
# Author: fadliwiryawirawan - https://twitter.com/11september1993
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.
#
# HOW TO:
# 1) Set bp at vbscript!CScriptRuntime::RunNoEH
# 2) Run
# 3) After bp is hit, launch script
# 4) ...

import sys, os, struct
import idc

class instr:
    def __init__(self, argsize, name, argfmt):
        self.argsize = argsize
        self.name = name
        self.argfmt = argfmt

def get_int8(pos):
    return struct.unpack("b", struct.pack("B", Byte(pos)))[0]

def get_uint8(pos):
    return Byte(pos)

def get_int16(pos):
    return struct.unpack("h", struct.pack("H", Word(pos)))[0]

def get_uint16(pos):
    return Word(pos)

def get_int32(pos):
    return struct.unpack("i", struct.pack("I", Dword(pos)))[0]

def get_uint32(pos):
    return Dword(pos)

def get_uint64(pos):
    return Qword(pos)

def get_type0(addr, pos):
    arg0 = get_uint8(pos)
    return " %d" % arg0

def get_type1(addr, pos):
    arg0 = get_uint16(pos)
    return " %d" % arg0

def get_type2(addr, pos):
    arg0 = get_int8(pos)
    return " %d" % arg0

def get_type3(addr, pos):
    arg0 = get_uint32(pos)
    return " %ld" % arg0

def get_type4(addr, pos):
    arg0 = get_int16(pos)
    return " %d" % arg0

def get_type5(addr, pos):
    arg0 = get_uint16(pos)
    return " %d" % arg0

def get_type6(addr, pos):
    arg0 = get_strlit_contents(addr + get_uint32(pos), -1, STRTYPE_C_16)
    return " '%s'" % arg0

def get_type7(addr, pos):
    arg0 = get_uint32(pos)
    return " %04X" % arg0

def get_type8(addr, pos):
    arg0 = get_uint64(pos)
    return " %.17g" % arg0

def get_type9(addr, pos):
    arg0 = get_int16(pos)
    arg1 = get_uint16(pos+2)
    return " %d %d" % (arg0, arg1)

def get_type10(addr, pos):
    arg0 = get_int16(pos)
    arg1 = get_int16(pos+2)
    return " %d %d" % (arg0, arg1)

def get_type11(addr, pos):
    arg0 = get_strlit_contents(addr + get_uint32(pos), -1, STRTYPE_C_16)
    arg1 = get_uint8(pos+4)
    if (arg1 == 0):
        return " '%s' FALSE" % arg0
    else:
        return " '%s' TRUE" % arg0

def get_type12(addr, pos):
    arg0 = get_strlit_contents(addr + get_uint32(pos), -1, STRTYPE_C_16)
    arg1 = get_uint16(pos+4)
    return " '%s' %d" % (arg0, arg1)

def get_type13(addr, pos):
    arg0 = get_int16(pos)
    arg1 = get_int16(pos+2)
    arg2 = get_int16(pos+4)
    return " %d %d %d" % (arg0, arg1, arg2)

def get_type14(addr, pos):
    arg0 = get_strlit_contents(addr + get_uint32(pos), -1, STRTYPE_C_16)
    arg1 = get_int16(pos+4)
    arg2 = get_int16(pos+6)
    return " '%s' %d %d" % (arg0, arg1, arg2)

def get_type15(addr, pos):
    arg0 = get_strlit_contents(addr + get_uint32(pos), -1, STRTYPE_C_16)
    arg1 = get_uint32(pos+4)
    arg2 = get_uint8(pos+8)
    if (arg2 == 0):
        return " '%s' %u FALSE" % (arg0, arg1)
    else:
        return " '%s' %u TRUE" % (arg0, arg1)

def get_type16(addr, pos):
    arg0 = get_strlit_contents(addr + get_uint32(pos), -1, STRTYPE_C_16)
    arg1 = get_int16(pos+4)
    arg2 = get_uint8(pos+6)
    if (arg2 == 0):
        return " '%s' %d FALSE" % (arg0, arg1)
    else:
        return " '%s' %d TRUE" % (arg0, arg1)

# Total = 0x6F
vbs_itable = {
    0x00: instr(0, "OP_None",           None),
    0x01: instr(0, "OP_FuncEnd",        None),
    0x02: instr(0, "OP_Bos0",           None),
    0x03: instr(1, "OP_Bos1",           get_type0),
    0x04: instr(2, "OP_Bos2",           get_type1),
    0x05: instr(4, "OP_Bos4",           get_type3),
    0x06: instr(4, "OP_DebugBreak",     get_type3),
    0x07: instr(4, "OP_ArrLclDim",      get_type9),
    0x08: instr(4, "OP_ArrLclReDim",    get_type9),
    0x09: instr(6, "OP_ArrNamDim",      get_type12),
    0x0A: instr(6, "OP_ArrNamReDim",    get_type12),
    0x0B: instr(1, "OP_IntConst",       get_type2), 
    0x0C: instr(4, "OP_LngConst",       get_type3),
    0x0D: instr(8, "OP_FltConst",       get_type8),
    0x0E: instr(4, "OP_StrConst",       get_type6),
    0x0F: instr(8, "OP_DateConst",      get_type8),
    0x10: instr(0, "OP_False",          None),
    0x11: instr(0, "OP_True",           None),
    0x12: instr(0, "OP_Null",           None),
    0x13: instr(0, "OP_Empty",          None),
    0x14: instr(0, "OP_NoArg",          None),
    0x15: instr(0, "OP_Nothing",        None),
    0x16: instr(5, "OP_ConstSt",        get_type11),
    0x17: instr(0, "OP_UNK_17",         None),
    0x18: instr(2, "OP_LocalLd",        get_type4),
    0x19: instr(2, "OP_LocalAdr",       get_type4),
    0x1A: instr(2, "OP_LocalSt",        get_type4),
    0x1B: instr(2, "OP_LocalSet",       get_type4),
    0x1C: instr(4, "OP_NamedLd",        get_type6),
    0x1D: instr(4, "OP_NamedAdr",       get_type6),
    0x1E: instr(4, "OP_NamedSt",        get_type6),
    0x1F: instr(4, "OP_NamedSet",       get_type6),
    0x20: instr(0, "OP_ThisLd",         None),
    0x21: instr(0, "OP_ThisSt",         None),
    0x22: instr(0, "OP_ThisSet",        None),
    0x23: instr(4, "OP_MemLd",          get_type6),
    0x24: instr(4, "OP_MemSt",          get_type6),
    0x25: instr(4, "OP_MemSet",         get_type6),
    0x26: instr(6, "OP_CallNmdLd",      get_type12),
    0x27: instr(6, "OP_CallNmdVoid",    get_type12), 
    0x28: instr(6, "OP_CallNmdAdr",     get_type12),
    0x29: instr(6, "OP_CallNmdSt",      get_type12),
    0x2A: instr(6, "OP_CallNmdSet",     get_type12),
    0x2B: instr(4, "OP_CallLclLd",      get_type9),
    0x2C: instr(4, "OP_CallLclVoid",    get_type9),
    0x2D: instr(4, "OP_CallLclAdr",     get_type9),
    0x2E: instr(4, "OP_CallLclSt",      get_type9),
    0x2F: instr(4, "OP_CallLclSet",     get_type9),
    0x30: instr(6, "OP_CallMemLd",      get_type12),
    0x31: instr(6, "OP_CallMemVoid",    get_type12),
    0x32: instr(6, "OP_CallMemSt",      get_type12),
    0x33: instr(6, "OP_CallMemSet",     get_type12),
    0x34: instr(2, "OP_CallIndLd",      get_type5),
    0x35: instr(2, "OP_CallIndVoid",    get_type5),
    0x36: instr(2, "OP_CallIndAdr",     get_type5),
    0x37: instr(2, "OP_CallIndSt",      get_type5),
    0x38: instr(2, "OP_CallIndSet",     get_type5),
    0x39: instr(0, "OP_Asg",            None),
    0x3A: instr(4, "OP_Jmp",            get_type7),
    0x3B: instr(4, "OP_JccTrue",        get_type7),
    0x3C: instr(4, "OP_JccFalse",       get_type7),
    0x3D: instr(0, "OP_Neg",            None),
    0x3E: instr(0, "OP_BitOr",          None),
    0x3F: instr(0, "OP_BitXor",         None),
    0x40: instr(0, "OP_BitAnd",         None),
    0x41: instr(0, "OP_BitNot",         None),
    0x42: instr(0, "OP_EQ",             None),
    0x43: instr(0, "OP_NE",             None), 
    0x44: instr(0, "OP_LT",             None), 
    0x45: instr(0, "OP_LE",             None), 
    0x46: instr(0, "OP_GT",             None), 
    0x47: instr(0, "OP_GE",             None), 
    0x48: instr(0, "OP_Add",            None),
    0x49: instr(0, "OP_Sub",            None),
    0x4A: instr(0, "OP_Mul",            None),
    0x4B: instr(0, "OP_Div",            None),
    0x4C: instr(0, "OP_Mod",            None),
    0x4D: instr(0, "OP_Eqv",            None),
    0x4E: instr(0, "OP_Pow",            None),
    0x4F: instr(0, "OP_Imp",            None),
    0x50: instr(0, "OP_Is",             None),
    0x51: instr(0, "OP_Like",           None),
    0x52: instr(0, "OP_Conc",           None),
    0x53: instr(0, "OP_Idiv",           None),
    0x54: instr(0, "OP_FixType",        None),
    0x55: instr(9, "OP_FnBind",         get_type15),
    0x56: instr(7, "OP_VarBind",        get_type16),
    0x57: instr(0, "OP_FnReturn",       None),
    0x58: instr(0, "OP_FnReturnEx",     None),
    0x59: instr(0, "OP_Pop",            None), 
    0x5A: instr(4, "OP_InitClass",      get_type6),
    0x5B: instr(4, "OP_CreateClass",    get_type6),
    0x5C: instr(9, "OP_FnBindEx",       get_type15),
    0x5D: instr(5, "OP_CreateVar",      get_type11),
    0x5E: instr(6, "OP_CreateArr",      get_type12),
    0x5F: instr(0, "OP_WithPush",       None),
    0x60: instr(0, "OP_WithPop",        None),
    0x61: instr(0, "OP_WithPop2",       None),
    0x62: instr(6, "OP_ForInitLocal",   get_type13),
    0x63: instr(6, "OP_ForNextLocal",   get_type13),
    0x64: instr(8, "OP_ForInitNamed",   get_type14), 
    0x65: instr(8, "OP_ForNextNamed",   get_type14),
    0x66: instr(1, "OP_OnError",        get_type2),
    0x67: instr(4, "OP_CaseEQ",         get_type7),
    0x68: instr(4, "OP_CaseNE",         get_type7),
    0x69: instr(6, "OP_ForInBegLcl",    get_type13),
    0x6A: instr(6, "OP_ForInNxtLcl",    get_type13),
    0x6B: instr(8, "OP_ForInBegNmd",    get_type14),
    0x6C: instr(8, "OP_ForInNxtNmd",    get_type14),
    0x6D: instr(4, "OP_ForInPop",       get_type10),
    0x6E: instr(0, "OP_UNK_6E",         None),
    0x6F: instr(0, "OP_UNK_6F",         None),
}

def pprint(indent_level, text):
    print("    " * indent_level + text)

def print_list(indent_level, addr, kind, count, list_ptr, start_id):

    pprint(indent_level, "%s count = %d" % (kind, count))

    if (list_ptr):

        indent_level += 1
    
        for i in range(count):

            name_pos = get_uint32(list_ptr + 8 * i)
            flags = get_uint32(list_ptr + 8 * i + 4)

            if (start_id):
                string = "%s %3d =" % (kind, start_id * (i + 1))
            else:
                string = "%s =" % kind
    
            if (start_id >= 0):
    
                if (start_id):
                    string += "    "
                
                elif (flags & 2):
                    string += " pri"
                
                else:
                    string += " pub"
                
            elif (flags & 0x200):
                string += " ref"
            
            else:
                string += " val"
            
            if (flags & 0x100):
                string += " Variant ()"
            else:
                string += " Variant   "
            
            string += " '%s'" % get_strlit_contents(addr + name_pos, -1, STRTYPE_C_16)

            pprint(indent_level, string)

def print_info(indent_level, addr, pfnc, func_id):

    name_pos = get_uint32(pfnc)
    stack = get_uint32(pfnc + 4)

    if (func_id):
        if (name_pos):
            name = get_strlit_contents(addr + name_pos, -1, STRTYPE_C_16)
            pprint(indent_level, "Function %d ('%s') [max stack = %u]:" % (func_id, name, stack))
        else:
            pprint(indent_level, "Function %d [max stack = %u]:" % (func_id, stack))
    else:
        pprint(indent_level, "Global code [max stack = %u]:" % stack)

    indent_level += 1

    flags = get_uint32(pfnc + 44)

    if (flags):

        string = "flags     = (%04lX)" % flags

        if (flags & 0x8000):
            string += " noconnect"

        if (flags & 0x4000):
            string += " sub"

        if (flags & 0x2000):
            string += " explicit"

        if (flags & 2):
            string += " private"

        pprint(indent_level, string)

    arg_count = get_int16(pfnc + 36)
    arg_addr = pfnc + 48

    if (func_id or arg_count > 0):
        print_list(indent_level, addr, "arg", arg_count, arg_addr, -1)

    lcl_count = get_int16(pfnc + 38)
    lcl_addr = arg_addr + 8 * arg_count

    if (func_id or lcl_count > 0):
        print_list(indent_level, addr, "lcl", lcl_count, lcl_addr, 1)

    tmp_count = get_int16(pfnc + 40)

    if (tmp_count > 0):
        print_list(indent_level, addr, "tmp", tmp_count, 0, 1)

def print_code(indent_level, addr, pfnc, bos_info, bos_data):

    code_start = addr + get_uint32(pfnc + 8)
    code_end = code_start + get_uint32(pfnc + 0xC)

    bos_info += get_uint32(pfnc + 0x10) * 8

    pprint(indent_level, "Pcode:")

    indent_level += 1

    position = code_start

    while (position < code_end):

        opcode = get_uint8(position)

        if opcode in (3,4,5):

            if (opcode == 3):
                value = get_uint8(position + 1)
    
            elif (opcode == 4):
                value = get_uint16(position + 1)
    
            elif (opcode == 5):
                value = get_uint32(position + 1)

            bos_info_pos = bos_info + 8 * value

            bos_start = get_uint32(bos_info_pos)
            bos_length = get_uint32(bos_info_pos + 4)

            string = "***BOS(%ld,%ld)***" % (bos_start, bos_start + bos_length) 

            if (bos_data):
                string += " %s *****" % get_strlit_contents(bos_data + 2 * bos_start, bos_length * 2, STRTYPE_C_16)

            pprint(indent_level, string)

            set_cmt(position, string, 0)

        string = "%04X    " % (position - code_start)

        if (opcode >= 0x6F):
            string += "ERROR(%u)" % opcode
            break

        string += "%-16s" % vbs_itable[opcode].name

        if (vbs_itable[opcode].argfmt is not None):
            string += vbs_itable[opcode].argfmt(addr, position + 1)

        pprint(indent_level, string)

        if opcode not in (3,4,5):
            set_cmt(position, string, 0)

        position += 1 + vbs_itable[opcode].argsize
  
def dump(addr):

    funcs = addr + get_uint32(addr + 0x10)
    funcs_count = get_uint32(addr + 0x14)
    bos_info = addr + get_uint32(addr + 0x1C)
    bos_data = addr + get_uint32(addr + 0x28)

    for i in range(funcs_count):

        pfnc = addr + get_uint32(funcs + i * 4)

        print_info(1, addr, pfnc, i)
        print_code(1, addr, pfnc, bos_info, bos_data)
        print

dump(Dword(get_reg_value("ECX")+0xC0))
