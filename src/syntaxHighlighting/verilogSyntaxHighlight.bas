Sub Syntax()
'
' Verilog Syntax Macro
'
'
Dim intruction_word()
' declare keywords, operators, etc for highlighting
instruction_word = Array("always", "and", "assign", "attribute", "begin", "buf", "bufif0", "bufif1", "case", "casex", "casez", "cmos", "deassign", "default", "defparam", "disable", "edge", "else", "end", "endattribute", "endcase", "endfunction", "endmodule", "endprimitive", "endspecify", "endtable", "endtask", "event", "for", "force", "forever", "fork", "function", "highz0", "highz1", "if", "ifnone", "initial", "inout", "input", "integer", "join", "medium", "module", "large", "localparam", "macromodule", "nand", "negedge", "nmos", _
 "nor", "not", "notif0", "notif1", "or", "output", "parameter", "pmos", "posedge", "primitive", "pull0", "pull1", "pulldown", "pullup", "rcmos", "real", "realtime", "reg", "release", "repeat", "rnmos", "rpmos", "rtran", "rtranif0", "rtranif1", "scalared", "signed", "small", "specify", "specparam", "strength", "strong0", "strong1", "supply0", "supply1", "table", "task", "time", "tran", "tranif0", "tranif1", "tri", "tri0", "tri1", "triand", "trior", "trireg", "unsigned", "vectored", "wait", _
 "wand", "weak0", "weak1", "while", "wire", "wor", "xnor", "xor", "alias", "always_comb", "always_ff", "always_latch", "assert", "assume", "automatic", "before", "bind", "bins", "binsof", "break", "constraint", "context", "continue", "cover", "cross", "design", "dist", "do", "expect", "export", "extends", "extern", "final", "first_match", "foreach", "forkjoin", "iff", "ignore_bins", "illegal_bins", "import", "incdir", "include", "inside", "instance", "intersect", "join_any", "join_none", "liblist", "library", "matches", _
 "modport", "new", "noshowcancelled", "null", "packed", "priority", "protected", "pulsestyle_onevent", "pulsestyle_ondetect", "pure", "rand", "randc", "randcase", "randsequence", "ref", "return", "showcancelled", "solve", "tagged", "this", "throughout", "timeprecision", "timeunit", "unique", "unique0", "use", "wait_order", "wildcard", "with", "within", "class", "clocking", "config", "generate", "covergroup", "interface", "package", "program", "property", "sequence", "endclass", "endclocking", "endconfig", "endgenerate", "endgroup", "endinterface", "endpackage", "endprogram", "endproperty", "endsequence", _
 "bit", "byte", "cell", "chandle", "const", "coverpoint", "enum", "genvar", "int", "local", "logic", "longint", "shortint", "shortreal", "static", "string", "struct", "super", "type", "typedef", "union", "var", "virtual", "void")
Dim keyword()
keyword = Array("$async$and$array", "$async$and$plane", "$async$nand$array", "$async$nand$plane", "$async$nor$array", "$async$nor$plane", "$async$or$array", "$async$or$plane", "$bitstoreal", "$countdrivers", "$display", "$displayb", "$displayh", "$displayo", "$dist_chi_square", "$dist_erlang", "$dist_exponential", "$dist_normal", "$dist_poisson", "$dist_t", "$dist_uniform", "$dumpall", "$dumpfile", "$dumpflush", "$dumplimit", "$dumpoff", "$dumpon", "$dumpportsall", "$dumpportsflush", "$dumpportslimit", "$dumpportsoff", "$dumpportson", "$dumpvars", "$fclose", "$fdisplayh", "$fdisplay", "$fdisplayf", "$fdisplayb", "$ferror", "$fflush", "$fgetc", "$fgets", "$finish", "$fmonitorb", "$fmonitor", "$fmonitorf", "$fmonitorh", "$fopen", "$fread", "$fscanf", _
 "$fseek", "$fsscanf", "$fstrobe", "$fstrobebb", "$fstrobef", "$fstrobeh", "$ftel", "$fullskew", "$fwriteb", "$fwritef", "$fwriteh", "$fwrite", "$getpattern", "$history", "$hold", "$incsave", "$input", "$itor", "$key", "$list", "$log", "$monitorb", "$monitorh", "$monitoroff", "$monitoron", "$monitor", "$monitoro", "$nochange", "$nokey", "$nolog", "$period", "$printtimescale", "$q_add", "$q_exam", "$q_full", "$q_initialize", "$q_remove", "$random", "$readmemb", "$readmemh", "$realtime", "$realtobits", "$recovery", "$recrem", "$removal", "$reset_count", "$reset", "$reset_value", "$restart", "$rewind", _
 "$rtoi", "$save", "$scale", "$scope", "$sdf_annotate", "$setup", "$setuphold", "$sformat", "$showscopes", "$showvariables", "$showvars", "$signed", "$skew", "$sreadmemb", "$sreadmemh", "$stime", "$stop", "$strobeb", "$strobe", "$strobeh", "$strobeo", "$swriteb", "$swriteh", "$swriteo", "$swrite", "$sync$and$array", "$sync$and$plane", "$sync$nand$array", "$sync$nand$plane", "$sync$nor$array", "$sync$nor$plane", "$sync$or$array", "$sync$or$plane", "$test$plusargs", "$time", "$timeformat", "$timeskew", "$ungetc", "$unsigned", "$value$plusargs", "$width", "$writeb", "$writeh", "$write", "$writeo")
Dim operators()
operators = Array("->", "+:", "-:", "#", "@", "@*", "=>", "*>", ",", ";", "{", "}", "+", "-", "/", "*", "**", "%", ">", ">=", ">>", ">>>", "<", "<=", "<<", "<<<", "!", "!=", "!==", "&", "&&", "|", "||", "=", "==", "===", "^", "^~", "~", "~^", "~&", "~|", "?", ":")
Dim preprocessor()
preprocessor = Array("`accelerate", "`autoexepand_vectornets", "`celldefine", "`default_nettype", "`define", "`default_decay_time", "`default_trireg_strength", "`delay_mode_distributed", "`delay_mode_path", "`delay_mode_unit", "`delay_mode_zero", "`else", "`elsif", "`endcelldefine", "`endif", "`endprotect", "`endprotected", "`expand_vectornets", "`file", "`ifdef", "`ifndef", "`include", "`line", "`noaccelerate", "`noexpand_vectornets", "`noremove_gatenames", "`noremove_netnames", "`nounconnected_drive", "`protect", "`protected", "`remove_gatenames", "`remove_netnames", "`resetall", "`timescale", "`unconnected_drive", "`undef", "`uselib")
Dim numtypes()
numtypes = Array("'b", "'B", "'o", "'O", "'d", "'D", "'h", "'H", "'sb", "'sB", "'so", "'sO", "'sd", "'sD", "'sh", "'sH", "'Sb", "'SB", "'So", "'SO", "'Sd", "'SD", "'Sh", "'SH", "'")

' disable spellcheck/proofreader in code segment
Selection.Range.LanguageID = wdEnglishUS
Selection.Range.NoProofing = True

For i = 1 To Selection.Paragraphs.Count
    Dim comm As Boolean
    comm = False
    Dim str As Boolean
    str = False
    Dim kword As Boolean
    kword = False
    For j = 1 To Selection.Paragraphs.Item(i).Range.Words.Count
        Dim myword As String
        myword = Selection.Paragraphs(i).Range.Words(j).Text
    If Strings.Asc(myword) < 33 Then GoTo ForEnd
        ''Debug.Print myword
        'highlights for comments
        If (Strings.InStr(1, myword, "//") > 0) Then
            comm = True
        End If
        If comm Then
            With Selection.Paragraphs(i).Range.Words(j).Font
             .TextColor.RGB = RGB(0, 128, 0)
            End With
        Else
            If Strings.InStr(2, myword, Chr(34)) > 0 And Not str Then
                str = True
            End If
            
            'highlights for strings
            If Strings.InStr(1, myword, Chr(34)) Then
                If Strings.InStr(1, myword, Chr(34)) = 1 Then
                    str = Not str
                Else
                    If Selection.Paragraphs(i).Range.Words(j).Characters(Strings.InStr(1, myword, Chr(34)) - 1) = Strings.Chr(92) Then
                        str = True
                    End If
                End If
                If Not (str) Then
                    With Selection.Paragraphs(i).Range.Words(j).Characters(1).Font
                     .TextColor.RGB = RGB(128, 128, 128)
                    End With
                End If
            End If
            If str Then
                With Selection.Paragraphs(i).Range.Words(j).Font
                 .TextColor.RGB = RGB(128, 128, 128)
                End With
            Else
                'highlights for keywords
                If kword Then
                    With Selection.Paragraphs(i).Range.Words(j).Font
                     .TextColor.RGB = RGB(128, 0, 255)
                    End With
                    kword = False
                Else
                    If Strings.InStr(1, myword, Chr(36)) = 1 Then
                        With Selection.Paragraphs(i).Range.Words(j).Characters(1).Font
                         .TextColor.RGB = RGB(128, 0, 255)
                        End With
                        kword = True
                    Else
                        'operators
                        Dim isOper As Boolean
                        isOper = False
                        For k = 0 To 42
                            If Strings.StrComp(operators(k), Trim(myword)) = 0 Then
                                isOper = True
                                k = 43
                            End If
                        Next
                        If isOper Then
                            With Selection.Paragraphs(i).Range.Words(j).Font
                             .Bold = True
                             .TextColor.RGB = RGB(0, 0, 128)
                            End With
                        Else
                        
                            'preprocessor
                            ''Debug.Print (myword)
                            If Strings.InStr(1, myword, Chr(96)) > 0 Then
                                With Selection.Paragraphs(i).Range.Words(j).Font
                                     .TextColor.RGB = RGB(128, 64, 0)
                                End With
                                With Selection.Paragraphs(i).Range.Words(j + 1).Font
                                     .TextColor.RGB = RGB(128, 64, 0)
                                End With
                                j = j + 1
                                While Selection.Paragraphs(i).Range.Words(j + 1).Characters(1) = Strings.Chr(95)
                                    With Selection.Paragraphs(i).Range.Words(j + 1).Font
                                         .TextColor.RGB = RGB(128, 64, 0)
                                    End With
                                    With Selection.Paragraphs(i).Range.Words(j + 2).Font
                                         .TextColor.RGB = RGB(128, 64, 0)
                                    End With
                                    j = j + 2
                                Wend
                            Else
                                If Strings.InStr(1, myword, Chr(39)) > 0 Or IsNumeric(myword) Then
                                    'numbers
                                    With Selection.Paragraphs(i).Range.Words(j).Font
                                     .TextColor.RGB = RGB(255, 128, 0)
                                    End With
                                Else
                                
                                    'highlights for instruction words - have this be lowest priority its slow af
                                    Dim commw As Boolean
                                    commw = False
                                    Dim multi As Boolean
                                    multi = False
                                    For k = 0 To 223
                                        While Selection.Paragraphs(i).Range.Words(j + 1).Characters(1) = Strings.Chr(95)
                                            myword = myword & "_" & Selection.Paragraphs(i).Range.Words(j + 2).Text
                                            ''Debug.Print myword & " final"
                                            j = j + 2
                                            multi = True
                                        Wend
                                        ''Debug.Print "enter"
                                        ''Debug.Print myword
                                        If Strings.StrComp(instruction_word(k), Trim(myword)) = 0 Then
                                        ''Debug.Print myword
                                        ''Debug.Print myword & " is equal to " & CStr(instruction_word(k)) & "!"
                                        commw = True
                                        k = 223
                                        End If
                                    Next
                                    If multi And commw Then
                                        With Selection.Paragraphs(i).Range.Words(j - 2).Font
                                         .Bold = True
                                         .TextColor.RGB = RGB(0, 0, 255)
                                        End With
                                        With Selection.Paragraphs(i).Range.Words(j - 1).Font
                                         .Bold = True
                                         .TextColor.RGB = RGB(0, 0, 255)
                                        End With
                                    End If
                                    If commw Then
                                        With Selection.Paragraphs(i).Range.Words(j).Font
                                         .Bold = True
                                         .TextColor.RGB = RGB(0, 0, 255)
                                        End With
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
ForEnd:
    Next
Next
End Sub