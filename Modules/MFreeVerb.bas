Attribute VB_Name = "MFreeVerb"
Option Explicit
'############################################################ v ' tuning.hpp ' v ############################################################'
'// Reverb model tuning values
'//
'// Written by Jezar at Dreampoint, June 2000
'// http://www.dreampoint.co.uk
'// This code is public domain
'
'#ifndef _tuning_
'#define _tuning_
'
'const int   numcombs        = 8;
'const int   numallpasses    = 4;
'const float muted           = 0;
'const float fixedgain       = 0.015f;

'const float scalewet        = 3;
'const float scaledry        = 2;
'const float scaledamp       = 0.4f;
'const float scaleroom       = 0.28f;
'const float offsetroom      = 0.7f;

'const float initialroom     = 0.5f;
'const float initialdamp     = 0.5f;
'const float initialwet      = 1/scalewet;
'const float initialdry      = 0;

'const float initialwidth    = 1;

'const float initialmode     = 0;
'const float freezemode      = 0.5f;
'const int   stereospread    = 23;

Const numcombs          As Long = 8
Const numallpasses      As Long = 4
Const muted           As Single = 0
Const fixedgain       As Single = 0.015

Const scalewet        As Single = 3
Const scaledry        As Single = 2
Const scaledamp       As Single = 0.4
Const scaleroom       As Single = 0.28
Const offsetroom      As Single = 0.7

Const initialroom     As Single = 0.5
Const initialdamp     As Single = 0.5
Const initialwet      As Single = 1 / scalewet
Const initialdry      As Single = 1

Const initialwidth    As Single = 1

Const initialmode     As Single = 0
Const freezemode      As Single = 0.5
Const stereospread      As Long = 23

'
'// These values assume 44.1KHz sample rate
'// they will probably be OK for 48KHz sample rate
'// but would need scaling for 96KHz (or other) sample rates.
'// The values were obtained by listening tests.
'const int combtuningL1      = 1116;
'const int combtuningR1      = 1116+stereospread;
'const int combtuningL2      = 1188;
'const int combtuningR2      = 1188+stereospread;
'const int combtuningL3      = 1277;
'const int combtuningR3      = 1277+stereospread;
'const int combtuningL4      = 1356;
'const int combtuningR4      = 1356+stereospread;
'const int combtuningL5      = 1422;
'const int combtuningR5      = 1422+stereospread;
'const int combtuningL6      = 1491;
'const int combtuningR6      = 1491+stereospread;
'const int combtuningL7      = 1557;
'const int combtuningR7      = 1557+stereospread;
'const int combtuningL8      = 1617;
'const int combtuningR8      = 1617+stereospread;
'const int allpasstuningL1   = 556;
'const int allpasstuningR1   = 556+stereospread;
'const int allpasstuningL2   = 441;
'const int allpasstuningR2   = 441+stereospread;
'const int allpasstuningL3   = 341;
'const int allpasstuningR3   = 341+stereospread;
'const int allpasstuningL4   = 225;
'const int allpasstuningR4   = 225+stereospread;

Const combtuningL1      As Long = 1116
Const combtuningR1      As Long = 1116 + stereospread
Const combtuningL2      As Long = 1188
Const combtuningR2      As Long = 1188 + stereospread
Const combtuningL3      As Long = 1277
Const combtuningR3      As Long = 1277 + stereospread
Const combtuningL4      As Long = 1356
Const combtuningR4      As Long = 1356 + stereospread
Const combtuningL5      As Long = 1422
Const combtuningR5      As Long = 1422 + stereospread
Const combtuningL6      As Long = 1491
Const combtuningR6      As Long = 1491 + stereospread
Const combtuningL7      As Long = 1557
Const combtuningR7      As Long = 1557 + stereospread
Const combtuningL8      As Long = 1617
Const combtuningR8      As Long = 1617 + stereospread
Const allpasstuningL1   As Long = 556
Const allpasstuningR1   As Long = 556 + stereospread
Const allpasstuningL2   As Long = 441
Const allpasstuningR2   As Long = 441 + stereospread
Const allpasstuningL3   As Long = 341
Const allpasstuningR3   As Long = 341 + stereospread
Const allpasstuningL4   As Long = 225
Const allpasstuningR4   As Long = 225 + stereospread


'
'#endif//_tuning_
'
'//ends
'############################################################ ^ ' tuning.hpp  ' ^ ############################################################'

'############################################################ v ' allpass.hpp ' v ############################################################'
'// Allpass filter declaration
'//
'// Written by Jezar at Dreampoint, June 2000
'// http://www.dreampoint.co.uk
'// This code is public domain
'
'#ifndef _allpass_
'#define _allpass_
'#include "denormals.h"
'
'Class allpass
'{
'public:
'                    allpass();
'            void    setbuffer(float *buf, int size);
'    inline  float   process(float inp);
'            void    mute();
'            void    setfeedback(float val);
'            float   getfeedback();
'// private:
'    float   feedback;
'    float   *buffer;
'    int     bufsize;
'    int     bufidx;
'};
Type allpass
    feedback As Single
    buffer() As Single
    bufsize  As Long
    bufidx   As Long
End Type

'############################################################ ^ ' allpass.hpp  ' ^ ############################################################'

'############################################################ v '   comb.hpp   ' v ############################################################'
'// Comb filter class declaration
'//
'// Written by Jezar at Dreampoint, June 2000
'// http://www.dreampoint.co.uk
'// This code is public domain
'
'#ifndef _comb_
'#define _comb_
'
'#include "denormals.h"
'
'Class comb
'{
'public:
'                    comb();
'            void    setbuffer(float *buf, int size);
'    inline  float   process(float inp);
'            void    mute();
'            void    setdamp(float val);
'            float   getdamp();
'            void    setfeedback(float val);
'            float   getfeedback();
'private:
'    float   feedback;
'    float   filterstore;
'    float   damp1;
'    float   damp2;
'    float   *buffer;
'    int     bufsize;
'    int     bufidx;
'};
Type comb
    feedback    As Single
    filterstore As Single
    damp1       As Single
    damp2       As Single
    buffer()    As Single
    bufsize     As Long
    bufidx      As Long
End Type

'############################################################ ^ '   comb.hpp   ' ^ ############################################################'

'############################################################ v ' revmodel.hpp ' v ############################################################'
'// Reverb model declaration
'//
'// Written by Jezar at Dreampoint, June 2000
'// http://www.dreampoint.co.uk
'// This code is public domain
'
'#ifndef _revmodel_
'#define _revmodel_
'
'#include "comb.hpp"
'#include "allpass.hpp"
'#include "tuning.h"
'
'Class revmodel
'{
'public:
'                    revmodel();
'            void    mute();
'            void    processmix(float *inputL, float *inputR, float *outputL, float *outputR, long numsamples, int skip);
'            void    processreplace(float *inputL, float *inputR, float *outputL, float *outputR, long numsamples, int skip);
'            void    setroomsize(float value);
'            float   getroomsize();
'            void    setdamp(float value);
'            float   getdamp();
'            void    setwet(float value);
'            float   getwet();
'            void    setdry(float value);
'            float   getdry();
'            void    setwidth(float value);
'            float   getwidth();
'            void    setmode(float value);
'            float   getmode();
'private:
'            void    update();
'private:
'    float   gain;
'    float   roomsize,roomsize1;
'    float   damp,damp1;
'    float   wet,wet1,wet2;
'    float   dry;
'    float   width;
'    float   mode;
'
'    // The following are all declared inline
'    // to remove the need for dynamic allocation
'    // with its subsequent error-checking messiness
'
'    // Comb filters
'    comb    combL[numcombs];
'    comb    combR[numcombs];
'
'    // Allpass filters
'    allpass allpassL[numallpasses];
'    allpass allpassR[numallpasses];
'
'    // Buffers for the combs
'    float   bufcombL1[combtuningL1];
'    float   bufcombR1[combtuningR1];
'    float   bufcombL2[combtuningL2];
'    float   bufcombR2[combtuningR2];
'    float   bufcombL3[combtuningL3];
'    float   bufcombR3[combtuningR3];
'    float   bufcombL4[combtuningL4];
'    float   bufcombR4[combtuningR4];
'    float   bufcombL5[combtuningL5];
'    float   bufcombR5[combtuningR5];
'    float   bufcombL6[combtuningL6];
'    float   bufcombR6[combtuningR6];
'    float   bufcombL7[combtuningL7];
'    float   bufcombR7[combtuningR7];
'    float   bufcombL8[combtuningL8];
'    float   bufcombR8[combtuningR8];
'
'    // Buffers for the allpasses
'    float   bufallpassL1[allpasstuningL1];
'    float   bufallpassR1[allpasstuningR1];
'    float   bufallpassL2[allpasstuningL2];
'    float   bufallpassR2[allpasstuningR2];
'    float   bufallpassL3[allpasstuningL3];
'    float   bufallpassR3[allpasstuningR3];
'    float   bufallpassL4[allpasstuningL4];
'    float   bufallpassR4[allpasstuningR4];
'};
'
'#endif//_revmodel_
'
'//ends
Type revmodel
    gain      As Single
    Roomsize  As Single
    roomsize1 As Single
    damp      As Single
    damp1     As Single
    
    wet       As Single
    wet1      As Single
    wet2      As Single
    
    dry       As Single
    Width     As Single
    mode      As Single

    '// The following are all declared inline
    '// to remove the need for dynamic allocation
    '// with its subsequent error-checking messiness
'das würde theoretisch auch in VB gehen, aber dann ist der Type zu groß, weil nur bis max 64k
    '// Comb filters
    combL() As comb 'numcombs = 8
    combR() As comb

    '// Allpass filters
    allpassL() As allpass 'numallpasses = 4
    allpassR() As allpass

    '// Buffers for the combs
    bufcombL1() As Single
    bufcombR1() As Single
    bufcombL2() As Single
    bufcombR2() As Single
    bufcombL3() As Single
    bufcombR3() As Single
    bufcombL4() As Single
    bufcombR4() As Single
    bufcombL5() As Single
    bufcombR5() As Single
    bufcombL6() As Single
    bufcombR6() As Single
    bufcombL7() As Single
    bufcombR7() As Single
    bufcombL8() As Single
    bufcombR8() As Single

    '// Buffers for the allpasses
    bufallpassL1() As Single
    bufallpassR1() As Single
    bufallpassL2() As Single
    bufallpassR2() As Single
    bufallpassL3() As Single
    bufallpassR3() As Single
    bufallpassL4() As Single
    bufallpassR4() As Single
End Type

Private isParamUpdating As Boolean
'############################################################ ^ ' revmodel.hpp ' ^ ############################################################'

'############################################################ v '  allpass.hpp ' v ############################################################'
'
'
'// Big to inline - but crucial for speed
'
'inline float allpass::process(float input)
'{
'    float output;
'    float bufout;
'
'    bufout = buffer[bufidx];
'    undenormalise(bufout);
'
'    output = -input + bufout;
'    buffer[bufidx] = input + (bufout*feedback);
'
'    if(++bufidx>=bufsize) bufidx = 0;
'
'    return output;
'}
'
'#endif//_allpass
'
'//ends
Function allpass_process(a As allpass, ByVal iinput As Single) As Single
    
    Dim output As Single
    Dim bufout As Single
    With a
        bufout = .buffer(.bufidx)
        undenormalise bufout
        
        output = -iinput + bufout
        .buffer(.bufidx) = iinput + (bufout * .feedback)
        
        .bufidx = .bufidx + 1
        If (.bufidx >= .bufsize) Then .bufidx = 0
    End With
    allpass_process = output
    
End Function
'############################################################ ^ ' allpass.hpp ' ^ ############################################################'

'############################################################ v ' allpass.cpp ' v ############################################################'
'void allpass::setbuffer(float *buf, int size)
'{
'    buffer = buf;
'    bufsize = size;
'}
Sub allpass_setbuffer(a As allpass, buf() As Single, ByVal size As Long)
    a.buffer = buf
    a.bufsize = size
End Sub

'void allpass::mute()
'{
'    for (int i=0; i<bufsize; i++)
'        buffer[i]=0;
'}
Sub allpass_mute(a As allpass)
    Dim i As Long
    For i = 0 To a.bufsize
        a.buffer(i) = 0
    Next
End Sub
'void allpass::setfeedback(float val)
'{
'    feedback = val;
'}
'float allpass::getfeedback()
'{
'    return feedback;
'}
Sub allpass_setfeedback(a As allpass, ByVal val As Single)
    a.feedback = val
End Sub
Function allpass_getfeedback(a As allpass) As Single
    allpass_getfeedback = a.feedback
End Function

'############################################################ ^ ' allpass.cpp ' ^ ############################################################'

'############################################################ v '   comb.hpp  ' v ############################################################'

'// Big to inline - but crucial for speed
'
'inline float comb::process(float input)
'{
'    float output;
'
'    output = buffer[bufidx];
'    undenormalise(output);
'
'    filterstore = (output*damp2) + (filterstore*damp1);
'    undenormalise(filterstore);
'
'    buffer[bufidx] = input + (filterstore*feedback);
'
'    if(++bufidx>=bufsize) bufidx = 0;
'
'    return output;
'}
'
'#endif //_comb_
'
'//ends
Function comb_process(c As comb, ByVal iinput As Single) As Single
    
    Dim output As Single
    With c
        output = .buffer(.bufidx)
        undenormalise output
        
        .filterstore = (output * .damp2) + (.filterstore * .damp1)
        undenormalise .filterstore
        
        .buffer(.bufidx) = iinput + (.filterstore * .feedback)
        .bufidx = .bufidx + 1
        If .bufidx >= .bufsize Then .bufidx = 0
    End With
    comb_process = output
    
End Function
'############################################################ ^ '  comb.hpp  ' ^ ############################################################'

'############################################################ v '  comb.cpp  ' v ############################################################'
'void comb::setbuffer(float *buf, int size)
'{
'    buffer = buf;
'    bufsize = size;
'}
Sub comb_setbuffer(c As comb, buf() As Single, ByVal size As Long)
    c.buffer = buf
    c.bufsize = size
End Sub
'
'void comb::mute()
'{
'    for (int i=0; i<bufsize; i++)
'        buffer[i]=0;
'}
Sub comb_mute(c As comb)
    Dim i As Long
    For i = 0 To c.bufsize
        c.buffer(i) = 0
    Next
End Sub

'
'void comb::setdamp(float val)
'{
'    damp1 = val;
'    damp2 = 1-val;
'}
Sub comb_setdamp(c As comb, ByVal val As Single)
    c.damp1 = val
    c.damp2 = 1 - val
End Sub

'float comb::getdamp()
'{
'    return damp1;
'}
Function comb_getdamp(c As comb) As Single
    comb_getdamp = c.damp1
End Function
'
'void comb::setfeedback(float val)
'{
'    feedback = val;
'}
'
'float comb::getfeedback()
'{
'    return feedback;
'}
Sub comb_setfeedback(c As comb, ByVal val As Single)
    c.feedback = val
End Sub
Function comb_getfeedback(c As comb) As Single
    comb_getfeedback = c.feedback
End Function
'############################################################ ^ '   comb.cpp   ' ^ ############################################################'

'############################################################ v ' revmodel.cpp ' v ############################################################'
'// Reverb model implementation
'//
'// Written by Jezar at Dreampoint, June 2000
'// http://www.dreampoint.co.uk
'// This code is public domain
'
'#include "revmodel.hpp"
'
'revmodel::revmodel()
'{
'    // Tie the components to their buffers
'    combL[0].setbuffer(bufcombL1,combtuningL1);
'    combR[0].setbuffer(bufcombR1,combtuningR1);
'    combL[1].setbuffer(bufcombL2,combtuningL2);
'    combR[1].setbuffer(bufcombR2,combtuningR2);
'    combL[2].setbuffer(bufcombL3,combtuningL3);
'    combR[2].setbuffer(bufcombR3,combtuningR3);
'    combL[3].setbuffer(bufcombL4,combtuningL4);
'    combR[3].setbuffer(bufcombR4,combtuningR4);
'    combL[4].setbuffer(bufcombL5,combtuningL5);
'    combR[4].setbuffer(bufcombR5,combtuningR5);
'    combL[5].setbuffer(bufcombL6,combtuningL6);
'    combR[5].setbuffer(bufcombR6,combtuningR6);
'    combL[6].setbuffer(bufcombL7,combtuningL7);
'    combR[6].setbuffer(bufcombR7,combtuningR7);
'    combL[7].setbuffer(bufcombL8,combtuningL8);
'    combR[7].setbuffer(bufcombR8,combtuningR8);
'    allpassL[0].setbuffer(bufallpassL1,allpasstuningL1);
'    allpassR[0].setbuffer(bufallpassR1,allpasstuningR1);
'    allpassL[1].setbuffer(bufallpassL2,allpasstuningL2);
'    allpassR[1].setbuffer(bufallpassR2,allpasstuningR2);
'    allpassL[2].setbuffer(bufallpassL3,allpasstuningL3);
'    allpassR[2].setbuffer(bufallpassR3,allpasstuningR3);
'    allpassL[3].setbuffer(bufallpassL4,allpasstuningL4);
'    allpassR[3].setbuffer(bufallpassR4,allpasstuningR4);
'
'    // Set default values
'    allpassL[0].setfeedback(0.5f);
'    allpassR[0].setfeedback(0.5f);
'    allpassL[1].setfeedback(0.5f);
'    allpassR[1].setfeedback(0.5f);
'    allpassL[2].setfeedback(0.5f);
'    allpassR[2].setfeedback(0.5f);
'    allpassL[3].setfeedback(0.5f);
'    allpassR[3].setfeedback(0.5f);
'    setwet(initialwet);
'    setroomsize(initialroom);
'    setdry(initialdry);
'    setdamp(initialdamp);
'    setwidth(initialwidth);
'    setmode(initialmode);
'
'    // Buffer will be full of rubbish - so we MUST mute them
'    mute();
'}
Function New_revmodel() As revmodel

    '// Tie the components to their buffers
    With New_revmodel
    
        ReDim .combL(0 To numcombs - 1) 'numcombs = 8
        ReDim .combR(0 To numcombs - 1)

    '// Allpass filters
        ReDim .allpassL(0 To numallpasses - 1) 'numallpasses = 4
        ReDim .allpassR(0 To numallpasses - 1)

    '// Buffers for the combs
        ReDim .bufcombL1(0 To combtuningL1 - 1)
        ReDim .bufcombR1(0 To combtuningR1 - 1)
        ReDim .bufcombL2(0 To combtuningL2 - 1)
        ReDim .bufcombR2(0 To combtuningR2 - 1)
        ReDim .bufcombL3(0 To combtuningL3 - 1)
        ReDim .bufcombR3(0 To combtuningR3 - 1)
        ReDim .bufcombL4(0 To combtuningL4 - 1)
        ReDim .bufcombR4(0 To combtuningR4 - 1)
        ReDim .bufcombL5(0 To combtuningL5 - 1)
        ReDim .bufcombR5(0 To combtuningR5 - 1)
        ReDim .bufcombL6(0 To combtuningL6 - 1)
        ReDim .bufcombR6(0 To combtuningR6 - 1)
        ReDim .bufcombL7(0 To combtuningL7 - 1)
        ReDim .bufcombR7(0 To combtuningR7 - 1)
        ReDim .bufcombL8(0 To combtuningL8 - 1)
        ReDim .bufcombR8(0 To combtuningR8 - 1)

    '// Buffers for the allpasses
        ReDim .bufallpassL1(0 To allpasstuningL1 - 1)
        ReDim .bufallpassR1(0 To allpasstuningR1 - 1)
        ReDim .bufallpassL2(0 To allpasstuningL2 - 1)
        ReDim .bufallpassR2(0 To allpasstuningR2 - 1)
        ReDim .bufallpassL3(0 To allpasstuningL3 - 1)
        ReDim .bufallpassR3(0 To allpasstuningR3 - 1)
        ReDim .bufallpassL4(0 To allpasstuningL4 - 1)
        ReDim .bufallpassR4(0 To allpasstuningR4 - 1)

    
        comb_setbuffer .combL(0), .bufcombL1, combtuningL1
        comb_setbuffer .combR(0), .bufcombR1, combtuningR1
        comb_setbuffer .combL(1), .bufcombL2, combtuningL2
        comb_setbuffer .combR(1), .bufcombR2, combtuningR2
        comb_setbuffer .combL(2), .bufcombL3, combtuningL3
        comb_setbuffer .combR(2), .bufcombR3, combtuningR3
        comb_setbuffer .combL(3), .bufcombL4, combtuningL4
        comb_setbuffer .combR(3), .bufcombR4, combtuningR4
        comb_setbuffer .combL(4), .bufcombL5, combtuningL5
        comb_setbuffer .combR(4), .bufcombR5, combtuningR5
        comb_setbuffer .combL(5), .bufcombL6, combtuningL6
        comb_setbuffer .combR(5), .bufcombR6, combtuningR6
        comb_setbuffer .combL(6), .bufcombL7, combtuningL7
        comb_setbuffer .combR(6), .bufcombR7, combtuningR7
        comb_setbuffer .combL(7), .bufcombL8, combtuningL8
        comb_setbuffer .combR(7), .bufcombR8, combtuningR8
        allpass_setbuffer .allpassL(0), .bufallpassL1, allpasstuningL1
        allpass_setbuffer .allpassR(0), .bufallpassR1, allpasstuningR1
        allpass_setbuffer .allpassL(1), .bufallpassL2, allpasstuningL2
        allpass_setbuffer .allpassR(1), .bufallpassR2, allpasstuningR2
        allpass_setbuffer .allpassL(2), .bufallpassL3, allpasstuningL3
        allpass_setbuffer .allpassR(2), .bufallpassR3, allpasstuningR3
        allpass_setbuffer .allpassL(3), .bufallpassL4, allpasstuningL4
        allpass_setbuffer .allpassR(3), .bufallpassR4, allpasstuningR4

        '// Set default values
        allpass_setfeedback .allpassL(0), 0.5
        allpass_setfeedback .allpassR(0), 0.5
        allpass_setfeedback .allpassL(1), 0.5
        allpass_setfeedback .allpassR(1), 0.5
        allpass_setfeedback .allpassL(2), 0.5
        allpass_setfeedback .allpassR(2), 0.5
        allpass_setfeedback .allpassL(3), 0.5
        allpass_setfeedback .allpassR(3), 0.5
        
        'Freeverb_setParameters New_revmodel, initialwet, initialroom, initialdry, initialdamp, initialwidth, initialmode
        
        revmodel_setwet New_revmodel, initialwet
        revmodel_setroomsize New_revmodel, initialroom
        revmodel_setdry New_revmodel, initialdry
        revmodel_setdamp New_revmodel, initialdamp
        revmodel_setwidth New_revmodel, initialwidth
        revmodel_setmode New_revmodel, initialmode
    End With
    '// Buffer will be full of rubbish - so we MUST mute them
    'braucht man eigentlich nicht:
    'revmodel_mute(rev)
End Function

Sub Freeverb_setParameters(rev As revmodel, ByVal aWet As Single, ByVal aRoomsize As Single, ByVal aDry As Single, ByVal aDamp As Single, ByVal aWidth As Single, ByVal aMode As Single)
    
    isParamUpdating = True
    
    revmodel_setwet rev, aWet
    revmodel_setroomsize rev, aRoomsize
    revmodel_setdry rev, aDry
    revmodel_setdamp rev, aDamp
    revmodel_setwidth rev, aWidth
    revmodel_setmode rev, aMode
    
    isParamUpdating = False
    MFreeVerb.revmodel_update rev
End Sub


'void revmodel::mute()
'{
'    if (getmode() >= freezemode)
'        return;
'
'    for (int i=0;i<numcombs;i++)
'    {
'        combL[i].mute();
'        combR[i].mute();
'    }
'    for (i=0;i<numallpasses;i++)
'    {
'        allpassL[i].mute();
'        allpassR[i].mute();
'    }
'}
Sub revmodel_mute(rev As revmodel)
    Dim i As Long
    With rev
        If (revmodel_getmode(rev) >= freezemode) Then _
            Exit Sub

        For i = 0 To numcombs - 1
        
            comb_mute .combL(i)
            comb_mute .combR(i)
            
        Next
        For i = 0 To numallpasses - 1
        
            allpass_mute .allpassL(i)
            allpass_mute .allpassR(i)
            
        Next
    End With
End Sub

'void revmodel::processreplace(float *inputL, float *inputR, float *outputL, float *outputR, long numsamples, int skip)
'{
'    float outL,outR,input;
'
'    while(numsamples-- > 0)
'    {
'        outL = outR = 0;
'        input = (*inputL + *inputR) * gain;
'
'        // Accumulate comb filters in parallel
'        for(int i=0; i<numcombs; i++)
'        {
'            outL += combL[i].process(input);
'            outR += combR[i].process(input);
'        }
'
'        // Feed through allpasses in series
'        for(i=0; i<numallpasses; i++)
'        {
'            outL = allpassL[i].process(outL);
'            outR = allpassR[i].process(outR);
'        }
'
'        // Calculate output REPLACING anything already there
'        *outputL = outL*wet1 + outR*wet2 + *inputL*dry;
'        *outputR = outR*wet1 + outL*wet2 + *inputR*dry;
'
'        // Increment sample pointers, allowing for interleave (if any)
'        inputL += skip;
'        inputR += skip;
'        outputL += skip;
'        outputR += skip;
'    }
'}
Sub revmodel_processreplace(rev As revmodel, ByRef inputL As Single, ByRef inputR As Single, ByRef outputL As Single, ByRef outputR As Single, ByVal numsamples As Long, ByVal skip As Long)

    Dim outL As Single, outR As Single, iinput As Single
    With rev
        'While (numsamples > 0)
            numsamples = numsamples - 1
            outL = 0
            outR = 0
            iinput = (inputL + inputR) * .gain
            Dim i As Long
            '// Accumulate comb filters in parallel
            For i = 0 To numcombs - 1
            
                outL = outL + comb_process(.combL(i), iinput)
                outR = outR + comb_process(.combR(i), iinput)
                
            Next

            '// Feed through allpasses in series
            For i = 0 To numallpasses - 1
            
                outL = allpass_process(.allpassL(i), outL)
                outR = allpass_process(.allpassR(i), outR)
                
            Next

            '// Calculate output REPLACING anything already there
            outputL = outL * .wet1 + outR * .wet2 + inputL * .dry
            outputR = outR * .wet1 + outL * .wet2 + inputR * .dry

            '// Increment sample pointers, allowing for interleave (if any)
            'inputL = inputL + skip
            'inputR = inputR + skip
            'outputL = outputL + skip
            'outputR = outputR + skip
        'Wend
    End With
End Sub


'void revmodel::processmix(float *inputL, float *inputR, float *outputL, float *outputR, long numsamples, int skip)
'{
'    float outL,outR,input;
'
'    while(numsamples-- > 0)
'    {
'        outL = outR = 0;
'        input = (*inputL + *inputR) * gain;
'
'        // Accumulate comb filters in parallel
'        for(int i=0; i<numcombs; i++)
'        {
'            outL += combL[i].process(input);
'            outR += combR[i].process(input);
'        }
'
'        // Feed through allpasses in series
'        for(i=0; i<numallpasses; i++)
'        {
'            outL = allpassL[i].process(outL);
'            outR = allpassR[i].process(outR);
'        }
'
'        // Calculate output MIXING with anything already there
'        *outputL += outL*wet1 + outR*wet2 + *inputL*dry;
'        *outputR += outR*wet1 + outL*wet2 + *inputR*dry;
'
'        // Increment sample pointers, allowing for interleave (if any)
'        inputL += skip;
'        inputR += skip;
'        outputL += skip;
'        outputR += skip;
'    }
'}
Sub revmodel_processmix(rev As revmodel, ByVal inputL As Single, ByVal inputR As Single, ByRef outputL As Single, ByRef outputR As Single)

    Dim outL As Single, outR As Single, iinput As Single
    Dim i As Long
    With rev
        'While (numsamples > 0)
            'numsamples = numsamples - 1
            'outL = outR = 0
            iinput = (inputL + inputR) / 2 '* .gain
            
            '// Accumulate comb filters in parallel
            For i = 0 To numcombs - 1
                
                outL = outL + comb_process(.combL(i), iinput)
                outR = outR + comb_process(.combR(i), iinput)
                
            Next
            
            '// Feed through allpasses in series
            For i = 0 To numallpasses - 1
                
                outL = allpass_process(.allpassL(i), outL)
                outR = allpass_process(.allpassR(i), outR)
                
            Next
            
            outL = outL / (numcombs + numallpasses)
            outR = outR / (numcombs + numallpasses)
            
            '// Calculate output MIXING with anything already there
            outputL = (outputL + outL * .wet1 + outR * .wet2 + inputL * .dry) / 4
            outputR = (outputR + outR * .wet1 + outL * .wet2 + inputR * .dry) / 4
            
            '// Increment sample pointers, allowing for interleave (if any)
            'inputL = inputL + skip
            'inputR = inputR + skip
            'outputL = outputL + skip
            'outputR = outputR + skip
        'Wend
    End With
End Sub


'void revmodel::update()
'{
'// Recalculate internal values after parameter change
'
'    int i;
'
'    wet1 = wet*(width/2 + 0.5f);
'    wet2 = wet*((1-width)/2);
'
'    if (mode >= freezemode)
'    {
'        roomsize1 = 1;
'        damp1 = 0;
'        gain = muted;
'    }
'    Else
'    {
'        roomsize1 = roomsize;
'        damp1 = damp;
'        gain = fixedgain;
'    }
'
'    for(i=0; i<numcombs; i++)
'    {
'        combL[i].setfeedback(roomsize1);
'        combR[i].setfeedback(roomsize1);
'    }
'
'    for(i=0; i<numcombs; i++)
'    {
'        combL[i].setdamp(damp1);
'        combR[i].setdamp(damp1);
'    }
'}
Sub revmodel_update(rev As revmodel)

'// Recalculate internal values after parameter change

    Dim i As Long
    With rev
        .wet1 = .wet * (.Width / 2 + 0.5)
        .wet2 = .wet * ((1 - .Width) / 2)

        If (.mode >= freezemode) Then
        
            .roomsize1 = 1
            .damp1 = 0
            .gain = muted
        
        Else
        
            .roomsize1 = .Roomsize
            .damp1 = .damp
            .gain = fixedgain
            
        End If

        For i = 0 To numcombs - 1
        
            comb_setfeedback .combL(i), .roomsize1
            comb_setfeedback .combR(i), .roomsize1
            
        Next

        For i = 0 To numcombs - 1
        
            comb_setdamp .combL(i), .damp1
            comb_setdamp .combR(i), .damp1
            
        Next
    End With
End Sub


'// The following get/set functions are not inlined, because
'// speed is never an issue when calling them, and also
'// because as you develop the reverb model, you may
'// wish to take dynamic action when they are called.
'
'void revmodel::setroomsize(float value)
'{
'    roomsize = (value*scaleroom) + offsetroom;
'    update();
'}
'float revmodel::getroomsize()
'{
'    return (roomsize-offsetroom)/scaleroom;
'}
Sub revmodel_setroomsize(rev As revmodel, ByVal value As Single)
    With rev
        .Roomsize = (value * scaleroom) + offsetroom
        If Not isParamUpdating Then revmodel_update rev
    End With
End Sub
Function revmodel_getroomsize(rev As revmodel) As Single
    With rev
        revmodel_getroomsize = (.Roomsize - offsetroom) / scaleroom
    End With
End Function



'void revmodel::setdamp(float value)
'{
'    damp = value*scaledamp;
'    update();
'}
'
'float revmodel::getdamp()
'{
'    return damp/scaledamp;
'}
Sub revmodel_setdamp(rev As revmodel, ByVal value As Single)
    With rev
        .damp = value * scaledamp
        If Not isParamUpdating Then revmodel_update rev
    End With
End Sub
Function revmodel_getdamp(rev As revmodel) As Single
    With rev
        revmodel_getdamp = .damp / scaledamp
    End With
End Function



'void revmodel::setwet(float value)
'{
'    wet = value*scalewet;
'    update();
'}
'
'float revmodel::getwet()
'{
'    return wet/scalewet;
'}
Sub revmodel_setwet(rev As revmodel, ByVal value As Single)
    With rev
        .wet = value * scalewet
        If Not isParamUpdating Then revmodel_update rev
    End With
End Sub
Function revmodel_getwet(rev As revmodel) As Single
    With rev
        revmodel_getwet = .wet / scalewet
    End With
End Function




'void revmodel::setdry(float value)
'{
'    dry = value*scaledry;
'}
'float revmodel::getdry()
'{
'    return dry/scaledry;
'}
Sub revmodel_setdry(rev As revmodel, ByVal value As Single)
    With rev
        .dry = value * scaledry
        If Not isParamUpdating Then revmodel_update rev
    End With
End Sub
Function revmodel_getdry(rev As revmodel) As Single
    With rev
        revmodel_getdry = .dry / scaledry
    End With
End Function


'void revmodel::setwidth(float value)
'{
'    width = value;
'    update();
'}
'float revmodel::getwidth()
'{
'    return width;
'}
Sub revmodel_setwidth(rev As revmodel, ByVal value As Single)
    With rev
        .Width = value
        If Not isParamUpdating Then revmodel_update rev
    End With
End Sub
Function revmodel_getwidth(rev As revmodel) As Single
    With rev
        revmodel_getwidth = .Width
    End With
End Function



'void revmodel::setmode(float value)
'{
'    mode = value;
'    update();
'}
'
'float revmodel::getmode()
'{
'    if (mode >= freezemode)
'        return 1;
'    Else
'        return 0;
'}
Sub revmodel_setmode(rev As revmodel, ByVal value As Single)
    With rev
        .mode = value
        If Not isParamUpdating Then revmodel_update rev
    End With
End Sub
Function revmodel_getmode(rev As revmodel) As Single
    With rev
        If .mode >= freezemode Then
            revmodel_getmode = 1
        Else
            revmodel_getmode = 0
        End If
    End With
End Function

'//ends

