Attribute VB_Name = "MGVerb"
Option Explicit
'Links:
'======
'http://stackoverflow.com/questions/5318989/reverb-algorithm
'http://www.soundonsound.com/sos/Oct01/articles/advancedreverb1.asp
'http://www.earlevel.com/main/1997/01/19/a-bit-about-reverb/
'https://ccrma.stanford.edu/~jos/pasp/
'http://freeverb3vst.osdn.jp/downloads.shtml
'https://github.com/swh/lv2/tree/master/gverb
'http://wiki.audacityteam.org/wiki/GVerb

'############################################################ v ' gverbdsp.h ' v ############################################################'
'/*
'        Copyright (C) 1999 Juhana Sadeharju
'                       kouhia at nic.funet.fi
'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'    You should have received a copy of the GNU General Public License
'    along with this program; if not, write to the Free Software
'    Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.
'    */
'#ifndef GVERBDSP_H
'#define GVERBDSP_H
'
'#include "../include/ladspa-util.h"
'
'typedef struct {
'  int size;
'  int idx;
'  float *buf;
'} ty_fixeddelay;
Type ty_fixeddelay
    size  As Long
    idx   As Long
    buf() As Single
End Type
'
'typedef struct {
'  int size;
'  float coeff;
'  int idx;
'  float *buf;
'} ty_diffuser;
Type ty_diffuser
    size  As Long
    coeff As Single
    idx   As Long
    buf() As Single
End Type
'
'typedef struct {
'  float damping;
'  float delay;
'} ty_damper;
Type ty_damper
    damping As Single
    delay   As Single
End Type
'############################################################ ^ ' gverbdsp.h ' ^ ############################################################'

'############################################################ ^ '  gverb.h   ' ^ ############################################################'
'/*
'        Copyright (C) 1999 Juhana Sadeharju
'                       kouhia at nic.funet.fi
'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'    You should have received a copy of the GNU General Public License
'    along with this program; if not, write to the Free Software
'    Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.
'    */
'
'#ifndef GVERB_H
'#define GVERB_H
'
'#include <stdlib.h>
'#include <math.h>
'#include <string.h>
'#include "gverbdsp.h"
'#include "gverb.h"
'#include "../include/ladspa-util.h"
'
'#define FDNORDER 4
Const FDNORDER As Long = 4
Const RAND_MAX As Single = 1

'
'typedef struct {
'  int rate;
'  float inputbandwidth;
'  float taillevel;
'  float earlylevel;
'  ty_damper *inputdamper;
'  float maxroomsize;
'  float roomsize;
'  float revtime;
'  float maxdelay;
'  float largestdelay;
'  ty_fixeddelay **fdndels;
'  float *fdngains;
'  int *fdnlens;
'  ty_damper **fdndamps;
'  float fdndamping;
'  ty_diffuser **ldifs;
'  ty_diffuser **rdifs;
'  ty_fixeddelay *tapdelay;
'  int *taps;
'  float *tapgains;
'  float *d;
'  float *u;
'  float *f;
'  double alpha;
'} ty_gverb;
Type ty_gverb
    Rate           As Long
    inputbandwidth As Single
    taillevel      As Single
    earlylevel     As Single
    inputdamper    As ty_damper
    maxroomsize    As Single
    Roomsize       As Single
    revtime        As Single
    maxdelay       As Single
    largestdelay   As Single
    fdndels(0 To FDNORDER - 1)  As ty_fixeddelay '**
    fdngains(0 To FDNORDER - 1) As Single
    fdnlens(0 To FDNORDER - 1)  As Long
    fdndamps(0 To FDNORDER - 1) As ty_damper '**
    fdndamping     As Single
    ldifs(0 To FDNORDER - 1)    As ty_diffuser '**
    rdifs(0 To FDNORDER - 1)    As ty_diffuser '**
    tapdelay                    As ty_fixeddelay '*
    taps(0 To FDNORDER - 1)     As Long
    tapgains(0 To FDNORDER - 1) As Single
    d(0 To FDNORDER - 1)        As Single
    u(0 To FDNORDER - 1)        As Single
    f(0 To FDNORDER - 1)        As Single
    alpha                       As Double
End Type
'
'ty_gverb *gverb_new(int, float, float, float, float, float, float, float, float);
'void gverb_free(ty_gverb *);
'void gverb_flush(ty_gverb *);
'static void gverb_do(ty_gverb *, float, float *, float *);
'static void gverb_set_roomsize(ty_gverb *, float);
'static void gverb_set_revtime(ty_gverb *, float);
'static void gverb_set_damping(ty_gverb *, float);
'static void gverb_set_inputbandwidth(ty_gverb *, float);
'static void gverb_set_earlylevel(ty_gverb *, float);
'static void gverb_set_taillevel(ty_gverb *, float);
'
'/*
' * This FDN reverb can be made smoother by setting matrix elements at the
' * diagonal and near of it to zero or nearly zero. By setting diagonals to zero
' * means we remove the effect of the parallel comb structure from the
' * reverberation.  A comb generates uniform impulse stream to the reverberation
' * impulse response, and thus it is not good. By setting near diagonal elements
' * to zero means we remove delay sequences having consequtive delays of the
' * similar lenths, when the delays are in sorted in length with respect to
' * matrix element index. The matrix described here could be generated by
' * differencing Rocchesso's circulant matrix at max diffuse value and at low
' * diffuse value (approaching parallel combs).
' *
' * Example 1:
' * Set a(k,k), for all k, equal to 0.
' *
' * Example 2:
' * Set a(k,k), a(k,k-1) and a(k,k+1) equal to 0.
' *
' * Example 3: The transition to zero gains could be smooth as well.
' * a(k,k-1) and a(k,k+1) could be 0.3, and a(k,k-2) and a(k,k+2) could
' * be 0.5, say.
' */
'
'static inline void gverb_fdnmatrix(float *a, float *b)
'{
'  const float dl0 = a[0], dl1 = a[1], dl2 = a[2], dl3 = a[3];
'
'  b[0] = 0.5f*(+dl0 + dl1 - dl2 - dl3);
'  b[1] = 0.5f*(+dl0 - dl1 - dl2 + dl3);
'  b[2] = 0.5f*(-dl0 + dl1 - dl2 + dl3);
'  b[3] = 0.5f*(+dl0 + dl1 + dl2 + dl3);
'}
Sub gverb_fdnmatrix(a() As Single, b() As Single)
    Dim dl0 As Single: dl0 = a(0)
    Dim dl1 As Single: dl1 = a(1)
    Dim dl2 As Single: dl2 = a(2)
    Dim dl3 As Single: dl3 = a(3)
    b(0) = 0.5 * (dl0 + dl1 - dl2 - dl3)
    b(1) = 0.5 * (dl0 - dl1 - dl2 + dl3)
    b(2) = 0.5 * (-dl0 + dl1 - dl2 + dl3)
    b(3) = 0.5 * (dl0 + dl1 + dl2 + dl3)
End Sub
'
'static inline void gverb_do(ty_gverb *p, float x, float *yl, float *yr)
'{
'  float z;
'  unsigned int i;
'  float lsum,rsum,sum,sign;
'
'  if (isnan(x) || fabsf(x) > 100000.0f) {
'    x = 0.0f;
'  }
'
'  z = damper_do(p->inputdamper, x);
'
'  z = diffuser_do(p->ldifs[0],z);
'
'  for(i = 0; i < FDNORDER; i++) {
'    p->u[i] = p->tapgains[i]*fixeddelay_read(p->tapdelay,p->taps[i]);
'  }
'  fixeddelay_write(p->tapdelay,z);
'
'  for(i = 0; i < FDNORDER; i++) {
'    p->d[i] = damper_do(p->fdndamps[i],
'            p->fdngains[i]*fixeddelay_read(p->fdndels[i],
'                               p->fdnlens[i]));
'  }
'
'  sum = 0.0f;
'  sign = 1.0f;
'  for(i = 0; i < FDNORDER; i++) {
'    sum += sign*(p->taillevel*p->d[i] + p->earlylevel*p->u[i]);
'    sign = -sign;
'  }
'  sum += x*p->earlylevel;
'  lsum = sum;
'  rsum = sum;
'
'  gverb_fdnmatrix(p->d,p->f);
'
'  for(i = 0; i < FDNORDER; i++) {
'    fixeddelay_write(p->fdndels[i],p->u[i]+p->f[i]);
'  }
'
'  lsum = diffuser_do(p->ldifs[1],lsum);
'  lsum = diffuser_do(p->ldifs[2],lsum);
'  lsum = diffuser_do(p->ldifs[3],lsum);
'  rsum = diffuser_do(p->rdifs[1],rsum);
'  rsum = diffuser_do(p->rdifs[2],rsum);
'  rsum = diffuser_do(p->rdifs[3],rsum);
'
'  *yl = lsum;
'  *yr = rsum;
'}
Sub gverb_do(p As ty_gverb, ByVal x As Single, ByRef yl_out As Single, ByRef yr_out As Single)

  Dim z As Double
  Dim i As Long
  Dim lsum As Double, rsum As Double, sum As Double, sign As Double

  'If (isnan(x) Or Abs(x) > 10000#) Then
  If (Abs(x) > 10000#) Then
    x = 0#
  End If

  z = damper_do(p.inputdamper, x)

  z = diffuser_do(p.ldifs(0), z)

  For i = 0 To FDNORDER - 1
    p.u(i) = p.tapgains(i) * fixeddelay_read(p.tapdelay, p.taps(i))
  Next
  fixeddelay_write p.tapdelay, z

  For i = 0 To FDNORDER - 1
    p.d(i) = damper_do(p.fdndamps(i), _
            p.fdngains(i) * fixeddelay_read(p.fdndels(i), _
                               p.fdnlens(i)))
  Next

  sum = 0#
  sign = 1#
  For i = 0 To FDNORDER - 1
    sum = sum + sign * (p.taillevel * p.d(i) + p.earlylevel * p.u(i))
    sign = -sign
  Next
  sum = sum + x * p.earlylevel
  lsum = sum
  rsum = sum

  gverb_fdnmatrix p.d, p.f

  For i = 0 To FDNORDER - 1
    fixeddelay_write p.fdndels(i), p.u(i) + p.f(i)
  Next

  lsum = diffuser_do(p.ldifs(1), lsum)
  lsum = diffuser_do(p.ldifs(2), lsum)
  lsum = diffuser_do(p.ldifs(3), lsum)
  rsum = diffuser_do(p.rdifs(1), rsum)
  rsum = diffuser_do(p.rdifs(2), rsum)
  rsum = diffuser_do(p.rdifs(3), rsum)

  yl_out = lsum
  yr_out = rsum
End Sub

'
'static inline void gverb_set_roomsize(ty_gverb *p, const float a)
'{
'  unsigned int i;
'
'  if (a <= 1.0 || isnan(a)) {
'    p->roomsize = 1.0;
'  } else {
'    p->roomsize = a;
'  }
'  p->largestdelay = p->rate * p->roomsize * 0.00294f;
'
'  p->fdnlens[0] = f_round(1.000000f*p->largestdelay);
'  p->fdnlens[1] = f_round(0.816490f*p->largestdelay);
'  p->fdnlens[2] = f_round(0.707100f*p->largestdelay);
'  p->fdnlens[3] = f_round(0.632450f*p->largestdelay);
'  for(i = 0; i < FDNORDER; i++) {
'    p->fdngains[i] = -powf((float)p->alpha, p->fdnlens[i]);
'  }
'
'  p->taps[0] = 5+f_round(0.410f*p->largestdelay);
'  p->taps[1] = 5+f_round(0.300f*p->largestdelay);
'  p->taps[2] = 5+f_round(0.155f*p->largestdelay);
'  p->taps[3] = 5+f_round(0.000f*p->largestdelay);
'
'  for(i = 0; i < FDNORDER; i++) {
'    p->tapgains[i] = powf((float)p->alpha, p->taps[i]);
'  }
'}
Sub gverb_set_roomsize(p As ty_gverb, ByVal a As Single)

    Dim i As Long 'unsigned int
    
    'If (a <= 1# Or isnan(a)) Then
    If (a <= 1#) Then
        p.Roomsize = 1#
    Else
        p.Roomsize = a
    End If
    p.largestdelay = p.Rate * p.Roomsize * 0.00294
    
    p.fdnlens(0) = Round(1# * p.largestdelay)
    p.fdnlens(1) = Round(0.81649 * p.largestdelay)
    p.fdnlens(2) = Round(0.7071 * p.largestdelay)
    p.fdnlens(3) = Round(0.63245 * p.largestdelay)
    For i = 0 To FDNORDER - 1
          p.fdngains(i) = -(p.alpha ^ p.fdnlens(i))
    Next
    
    p.taps(0) = 5 + Round(0.41 * p.largestdelay)
    p.taps(1) = 5 + Round(0.3 * p.largestdelay)
    p.taps(2) = 5 + Round(0.155 * p.largestdelay)
    p.taps(3) = 5 + Round(0# * p.largestdelay)
    
    For i = 0 To FDNORDER - 1
        p.tapgains(i) = (p.alpha ^ p.taps(i))
    Next

End Sub

'
'static inline void gverb_set_revtime(ty_gverb *p,float a)
'{
'  float ga,gt;
'  double n;
'  unsigned int i;
'
'  p->revtime = a;
'
'  ga = 60.0;
'  gt = p->revtime;
'  ga = powf(10.0f,-ga/20.0f);
'  n = p->rate*gt;
'  p->alpha = (double)powf(ga,1.0f/n);
'
'  for(i = 0; i < FDNORDER; i++) {
'    p->fdngains[i] = -powf((float)p->alpha, p->fdnlens[i]);
'  }
'
'}
Sub gverb_set_revtime(p As ty_gverb, ByVal a As Single)

    Dim ga As Single, gt As Single
    Dim n As Double
    Dim i As Long
    
    p.revtime = a
    
    ga = 60#
    gt = p.revtime
    ga = 1# ^ (-ga / 2#)
    n = p.Rate * gt
    p.alpha = CDbl(ga ^ (1# / n))
    
    For i = 0 To FDNORDER - 1
        p.fdngains(i) = -(p.alpha ^ p.fdnlens(i))
    Next
End Sub

'
'static inline void gverb_set_damping(ty_gverb *p,float a)
'{
'  unsigned int i;
'
'  p->fdndamping = a;
'  for(i = 0; i < FDNORDER; i++) {
'    damper_set(p->fdndamps[i],p->fdndamping);
'  }
'}
Sub gverb_set_damping(p As ty_gverb, ByVal a As Single)

    Dim i As Long
    
    p.fdndamping = a
    For i = 0 To FDNORDER - 1
        damper_set p.fdndamps(i), p.fdndamping
    Next
End Sub

'
'static inline void gverb_set_inputbandwidth(ty_gverb *p,float a)
'{
'  p->inputbandwidth = a;
'  damper_set(p->inputdamper,1.0 - p->inputbandwidth);
'}
Sub gverb_set_inputbandwidth(p As ty_gverb, ByVal a As Single)

    p.inputbandwidth = a
    damper_set p.inputdamper, 1# - p.inputbandwidth
End Sub

'
'static inline void gverb_set_earlylevel(ty_gverb *p,float a)
'{
'  p->earlylevel = a;
'}
Sub gverb_set_earlylevel(p As ty_gverb, ByVal a As Single)
  
    p.earlylevel = a

End Sub


'
'static inline void gverb_set_taillevel(ty_gverb *p,float a)
'{
'  p->taillevel = a;
'}
Sub gverb_set_taillevel(p As ty_gverb, ByVal a As Single)
  
    p.taillevel = a

End Sub

'
'#End If
'############################################################ ^ '  gverb.h   ' ^ ############################################################'

'############################################################ v ' gverbdsp.c ' v ############################################################'

'
'ty_diffuser *diffuser_make(int, float);
'void diffuser_free(ty_diffuser *);
'void diffuser_flush(ty_diffuser *);
'//float diffuser_do(ty_diffuser *, float);
'
'ty_damper *damper_make(float);
'void damper_free(ty_damper *);
'void damper_flush(ty_damper *);
'//void damper_set(ty_damper *, float);
'//float damper_do(ty_damper *, float);
'
'ty_fixeddelay *fixeddelay_make(int);
'void fixeddelay_free(ty_fixeddelay *);
'void fixeddelay_flush(ty_fixeddelay *);
'//float fixeddelay_read(ty_fixeddelay *, int);
'//void fixeddelay_write(ty_fixeddelay *, float);
'
'int isprime(int);
'int nearest_prime(int, float);
'
'static inline float diffuser_do(ty_diffuser *p, float x)
'{
'  float y,w;
'
'  w = x - p->buf[p->idx]*p->coeff;
'  w = flush_to_zero(w);
'  y = p->buf[p->idx] + w*p->coeff;
'  p->buf[p->idx] = w;
'  p->idx = (p->idx + 1) % p->size;
'  return(y);
'}

Function diffuser_do(p As ty_diffuser, ByVal x As Single) As Single
    Dim y As Single, w As Single
    With p
        w = x - .buf(.idx) * .coeff
        w = flush_to_zero(w) '????
        'w = flush_to_zero(w)
        y = .buf(.idx) + w * .coeff
        .buf(.idx) = w
        .idx = (.idx + 1) Mod .size
    End With
    diffuser_do = y
End Function
'
'static inline float fixeddelay_read(ty_fixeddelay *p, int n)
'{
'  int i;
'
'  i = (p->idx - n + p->size) % p->size;
'  return(p->buf[i]);
'}
Function fixeddelay_read(p As ty_fixeddelay, ByVal n As Long) As Single
    Dim i As Long
    With p
        i = (.idx - n + .size) Mod .size
        fixeddelay_read = .buf(i)
    End With
End Function
'
'static inline void fixeddelay_write(ty_fixeddelay *p, float x)
'{
'  p->buf[p->idx] = x;
'  p->idx = (p->idx + 1) % p->size;
'}
Sub fixeddelay_write(p As ty_fixeddelay, ByVal x As Single)
    With p
        .buf(p.idx) = x
        .idx = (.idx + 1) Mod .size
    End With
End Sub

'
'static inline void damper_set(ty_damper *p, float damping)
'{
'  p->damping = damping;
'}
Sub damper_set(p As ty_damper, ByVal damping As Single)
  
    p.damping = damping

End Sub

'
'static inline float damper_do(ty_damper *p, float x)
'{
'  float y;
'
'  y = x*(1.0-p->damping) + p->delay*p->damping;
'  p->delay = y;
'  return(y);
'}
Function damper_do(p As ty_damper, ByVal x As Single) As Single
    Dim y As Single
    With p
        y = x * (1# - .damping) + .delay * .damping
        .delay = y
        damper_do = y
    End With
End Function
'
'#End If

''##########' gverbdsp.c '##########'

'/*
'        Copyright (C) 1999 Juhana Sadeharju
'                       kouhia at nic.funet.fi
'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'    You should have received a copy of the GNU General Public License
'    along with this program; if not, write to the Free Software
'    Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.
'    */
'
'#include <stdio.h>
'#include <math.h>
'#include <stdlib.h>
'#include <string.h>
'
'#include "gverbdsp.h"
'
'#define TRUE 1
'#define FALSE 0
'
'ty_diffuser *diffuser_make(int size, float coeff)
'{
'  ty_diffuser *p;
'  int i;
'
'  p = (ty_diffuser *)malloc(sizeof(ty_diffuser));
'  p->size = size;
'  p->coeff = coeff;
'  p->idx = 0;
'  p->buf = (float *)malloc(size*sizeof(float));
'  for (i = 0; i < size; i++) p->buf[i] = 0.0;
'  return(p);
'}
Function diffuser_make(ByVal size As Long, ByVal coeff As Single) As ty_diffuser
    Dim i As Long
    With diffuser_make
        .size = size
        .coeff = coeff
        '.idx = 0
        ReDim .buf(0 To size - 1)
    End With
End Function
'
'void diffuser_free(ty_diffuser * p)
'{
'  free(p->buf);
'  free(p);
'}
Sub diffuser_free(p As ty_diffuser)
    '
End Sub
'
'void diffuser_flush(ty_diffuser * p)
'{
'  memset(p->buf, 0, p->size * sizeof(float));
'}
Sub diffuser_flush(p As ty_diffuser)
    'RtlZeroMemory?
End Sub

'
'ty_damper *damper_make(float damping)
'{
'  ty_damper *p;
'
'  p = (ty_damper *)malloc(sizeof(ty_damper));
'  p->damping = damping;
'  p->delay = 0.0f;
'  return(p);
'}
Function damper_make(ByVal damping As Single) As ty_damper
    With damper_make
        .damping = damping
    End With
End Function
'
'void damper_free(ty_damper * p)
'{
'  free(p);
'}
Sub damper_free(p As ty_damper)
    '
End Sub
'
'void damper_flush(ty_damper * p)
'{
'  p->delay = 0.0f;
'}
Sub damper_flush(p As ty_damper)
    '
End Sub

'
'ty_fixeddelay *fixeddelay_make(int size)
'{
'  ty_fixeddelay *p;
'  int i;
'
'  p = (ty_fixeddelay *)malloc(sizeof(ty_fixeddelay));
'  p->size = size;
'  p->idx = 0;
'  p->buf = (float *)malloc(size*sizeof(float));
'  for (i = 0; i < size; i++) p->buf[i] = 0.0;
'  return(p);
'}
Function fixeddelay_make(ByVal size As Long) As ty_fixeddelay
    Dim i As Long
    With fixeddelay_make
        .size = size
        .idx = 0
        ReDim .buf(0 To .size - 1)
    End With
End Function
'
'void fixeddelay_free(ty_fixeddelay * p)
'{
'  free(p->buf);
'  free(p);
'}
Sub fixeddelay_free(p As ty_fixeddelay)
    '
End Sub
'
'void fixeddelay_flush(ty_fixeddelay * p)
'{
'  memset(p->buf, 0, p->size * sizeof(float));
'}
Sub fixeddelay_flush(p As ty_fixeddelay)
    'rtlzeromemory '???
End Sub
'
'int isprime(int n)
'{
'  unsigned int i;
'  const unsigned int lim = (int)sqrtf((float)n);
'
'  if (n == 2) return(TRUE);
'  if ((n & 1) == 0) return(FALSE);
'  for(i = 3; i <= lim; i += 2)
'    if ((n % i) == 0) return(FALSE);
'  return(TRUE);
'}
Function IsPrime(ByVal n As Long) As Boolean
    Dim i As Long
    Dim lim As Long: lim = Sqr(n)
    If n = 2 Then IsPrime = True: Exit Function
    If (n And 1) = 0 Then IsPrime = False: Exit Function
    For i = 3 To lim Step 2
        If (n Mod i) = 0 Then IsPrime = False: Exit Function
    Next
    IsPrime = True
End Function
'
'int nearest_prime(int n, float rerror)
'     /* relative error; new prime will be in range
'      * [n-n*rerror, n+n*rerror];
'      */
'{
'  int bound,k;
'
'  if (isprime(n)) return(n);
'  /* assume n is large enough and n*rerror enough smaller than n */
'  bound = n*rerror;
'  for(k = 1; k <= bound; k++) {
'    if (isprime(n+k)) return(n+k);
'    if (isprime(n-k)) return(n-k);
'  }
'  return(-1);
'}
Function nearest_prime(ByVal n As Long, ByVal rerror As Single) As Long
    'Debug.Print "nearest_prime"
    Dim k As Long, bound As Long
    If IsPrime(n) Then nearest_prime = n: Exit Function
    bound = n * rerror
    For k = 1 To bound
        If IsPrime(n + k) Then nearest_prime = n + k: Exit Function
        If IsPrime(n - k) Then nearest_prime = n - k: Exit Function
    Next
End Function
'############################################################ ^ ' gverbdsp.c ' ^ ############################################################'

'############################################################ v '  gverb.c   ' v ############################################################'
'
'/*
'        Copyright (C) 1999 Juhana Sadeharju
'                       kouhia at nic.funet.fi
'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'    You should have received a copy of the GNU General Public License
'    along with this program; if not, write to the Free Software
'    Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.
'    */
'
'
'#include <stdio.h>
'#include <stdlib.h>
'#include <math.h>
'#include <string.h>
'#include "gverbdsp.h"
'#include "gverb.h"
'#include "../include/ladspa-util.h"
'
'ty_gverb *gverb_new(int srate, float maxroomsize, float roomsize,
'            float revtime,
'            float damping, float spread,
'            float inputbandwidth, float earlylevel,
'            float taillevel)
'{
'  ty_gverb *p;
'  float ga,gb,gt;
'  int i,n;
'  float r;
'  float diffscale;
'  int a,b,c,cc,d,dd,e;
'  float spread1,spread2;
'
'  p = (ty_gverb *)malloc(sizeof(ty_gverb));
'  p->rate = srate;
'  p->fdndamping = damping;
'  p->maxroomsize = maxroomsize;
'  p->roomsize = roomsize;
'  p->revtime = revtime;
'  p->earlylevel = earlylevel;
'  p->taillevel = taillevel;
'
'  p->maxdelay = p->rate*p->maxroomsize/340.0;
'  p->largestdelay = p->rate*p->roomsize/340.0;
'
'
'  /* Input damper */
'
'  p->inputbandwidth = inputbandwidth;
'  p->inputdamper = damper_make(1.0 - p->inputbandwidth);
'
'
'  /* FDN section */
'
'
'  p->fdndels = (ty_fixeddelay **)calloc(FDNORDER, sizeof(ty_fixeddelay *));
'  for(i = 0; i < FDNORDER; i++) {
'    p->fdndels[i] = fixeddelay_make((int)p->maxdelay+1000);
'  }
'  p->fdngains = (float *)calloc(FDNORDER, sizeof(float));
'  p->fdnlens = (int *)calloc(FDNORDER, sizeof(int));
'
'  p->fdndamps = (ty_damper **)calloc(FDNORDER, sizeof(ty_damper *));
'  for(i = 0; i < FDNORDER; i++) {
'    p->fdndamps[i] = damper_make(p->fdndamping);
'  }
'
'  ga = 60.0;
'  gt = p->revtime;
'  ga = powf(10.0f,-ga/20.0f);
'  n = p->rate*gt;
'  p->alpha = pow((double)ga, 1.0/(double)n);
'
'  gb = 0.0;
'  for(i = 0; i < FDNORDER; i++) {
'    if (i == 0) gb = 1.000000*p->largestdelay;
'    if (i == 1) gb = 0.816490*p->largestdelay;
'    if (i == 2) gb = 0.707100*p->largestdelay;
'    if (i == 3) gb = 0.632450*p->largestdelay;
'
'#if 0
'    p->fdnlens[i] = nearest_prime((int)gb, 0.5);
'#Else
'    p->fdnlens[i] = f_round(gb);
'#End If
'    p->fdngains[i] = -powf((float)p->alpha,p->fdnlens[i]);
'  }
'
'  p->d = (float *)calloc(FDNORDER, sizeof(float));
'  p->u = (float *)calloc(FDNORDER, sizeof(float));
'  p->f = (float *)calloc(FDNORDER, sizeof(float));
'
'  /* Diffuser section */
'
'  diffscale = (float)p->fdnlens[3]/(210+159+562+410);
'  spread1 = spread;
'  spread2 = 3.0*spread;
'
'  b = 210;
'  r = 0.125541;
'  a = spread1*r;
'  c = 210+159+a;
'  cc = c-b;
'  r = 0.854046;
'  a = spread2*r;
'  d = 210+159+562+a;
'  dd = d-c;
'  e = 1341-d;
'
'  p->ldifs = (ty_diffuser **)calloc(4, sizeof(ty_diffuser *));
'  p->ldifs[0] = diffuser_make((int)(diffscale*b),0.75);
'  p->ldifs[1] = diffuser_make((int)(diffscale*cc),0.75);
'  p->ldifs[2] = diffuser_make((int)(diffscale*dd),0.625);
'  p->ldifs[3] = diffuser_make((int)(diffscale*e),0.625);
'
'  b = 210;
'  r = -0.568366;
'  a = spread1*r;
'  c = 210+159+a;
'  cc = c-b;
'  r = -0.126815;
'  a = spread2*r;
'  d = 210+159+562+a;
'  dd = d-c;
'  e = 1341-d;
'
'  p->rdifs = (ty_diffuser **)calloc(4, sizeof(ty_diffuser *));
'  p->rdifs[0] = diffuser_make((int)(diffscale*b),0.75);
'  p->rdifs[1] = diffuser_make((int)(diffscale*cc),0.75);
'  p->rdifs[2] = diffuser_make((int)(diffscale*dd),0.625);
'  p->rdifs[3] = diffuser_make((int)(diffscale*e),0.625);
'
'
'
'  /* Tapped delay section */
'
'  p->tapdelay = fixeddelay_make(44000);
'  p->taps = (int *)calloc(FDNORDER, sizeof(int));
'  p->tapgains = (float *)calloc(FDNORDER, sizeof(float));
'
'  p->taps[0] = 5+0.410*p->largestdelay;
'  p->taps[1] = 5+0.300*p->largestdelay;
'  p->taps[2] = 5+0.155*p->largestdelay;
'  p->taps[3] = 5+0.000*p->largestdelay;
'
'  for(i = 0; i < FDNORDER; i++) {
'    p->tapgains[i] = pow(p->alpha,(double)p->taps[i]);
'  }
'
'  return(p);
'}

Function gverb_new(ByVal srate As Long, ByVal maxroomsize As Single, ByVal Roomsize As Single, _
                   ByVal revtime As Single, _
                   ByVal damping As Single, ByVal spread As Single, _
                   ByVal inputbandwidth As Single, ByVal earlylevel As Single, _
                   ByVal taillevel As Single) As ty_gverb

  Dim p As ty_gverb
  Dim ga As Single, gb As Single, gt As Single
  Dim i As Long, n As Long
  Dim r As Single
  Dim diffscale As Single
  Dim a As Long, b As Long, c As Long, cc As Long, d As Long, dd As Long, e As Long
  Dim spread1 As Single, spread2 As Single

  'p = (ty_gverb *)malloc(sizeof(ty_gverb))
  p.Rate = srate
  p.fdndamping = damping
  p.maxroomsize = maxroomsize
  p.Roomsize = Roomsize
  p.revtime = revtime
  p.earlylevel = earlylevel
  p.taillevel = taillevel

  p.maxdelay = p.Rate * p.maxroomsize / 340#
  p.largestdelay = p.Rate * p.Roomsize / 340#


  '/* Input damper */

  p.inputbandwidth = inputbandwidth
  p.inputdamper = damper_make(1# - p.inputbandwidth)


  '/* FDN section */


  'p.fdndels = (ty_fixeddelay **)calloc(FDNORDER, sizeof(ty_fixeddelay *))
  'ReDim p.fdndels(0 To FDNORDER - 1)
  For i = 0 To FDNORDER - 1
    p.fdndels(i) = fixeddelay_make(p.maxdelay + 1000)
  Next
  'p.fdngains = (float *)calloc(FDNORDER, sizeof(float))
  'ReDim p.fdngains(0 To FDNORDER - 1)
  'p.fdnlens = (int *)calloc(FDNORDER, sizeof(int))
  'ReDim p.fdnlens(0 To FDNORDER - 1)

  'p.fdndamps = (ty_damper **)calloc(FDNORDER, sizeof(ty_damper *))
  'ReDim p.fdndamps(0 To FDNORDER - 1)
  For i = 0 To FDNORDER - 1
    p.fdndamps(i) = damper_make(p.fdndamping)
  Next

  ga = 60#
  gt = p.revtime
  ga = (1# ^ -ga / 2#)
  n = p.Rate * gt
  p.alpha = (ga ^ (1# / n))

  gb = 0#
  For i = 0 To FDNORDER - 1
    If (i = 0) Then gb = 1# * p.largestdelay
    If (i = 1) Then gb = 0.81649 * p.largestdelay
    If (i = 2) Then gb = 0.7071 * p.largestdelay
    If (i = 3) Then gb = 0.63245 * p.largestdelay

'#if 0
    p.fdnlens(i) = nearest_prime(gb, 0.5)
'#Else
'    p.fdnlens(i) = f_round(gb)
'#End If
    p.fdngains(i) = -CSng(p.alpha ^ p.fdnlens(i))
  Next

  'p.d = (float *)calloc(FDNORDER, sizeof(float))
  'ReDim p.d(0 To FDNORDER - 1)
  'p.u = (float *)calloc(FDNORDER, sizeof(float))
  'ReDim p.u(0 To FDNORDER - 1)
  'p.f = (float *)calloc(FDNORDER, sizeof(float))
  'ReDim p.f(0 To FDNORDER - 1)

  '/* Diffuser section */

  diffscale = CSng(p.fdnlens(3) / (210 + 159 + 562 + 410))
  spread1 = spread
  spread2 = 3# * spread

  b = 210
  r = 0.125541
  a = spread1 * r
  c = 210 + 159 + a
  cc = c - b
  r = 0.854046
  a = spread2 * r
  d = 210 + 159 + 562 + a
  dd = d - c
  e = 1341 - d

  'p.ldifs = (ty_diffuser **)calloc(4, sizeof(ty_diffuser *))
  'ReDim p.ldifs(0 To 3)
  p.ldifs(0) = diffuser_make((diffscale * b), 0.75)
  p.ldifs(1) = diffuser_make((diffscale * cc), 0.75)
  p.ldifs(2) = diffuser_make((diffscale * dd), 0.625)
  p.ldifs(3) = diffuser_make((diffscale * e), 0.625)

  b = 210
  r = -0.568366
  a = spread1 * r
  c = 210 + 159 + a
  cc = c - b
  r = -0.126815
  a = spread2 * r
  d = 210 + 159 + 562 + a
  dd = d - c
  e = 1341 - d

  'p.rdifs = (ty_diffuser **)calloc(4, sizeof(ty_diffuser *))
  'ReDim p.rdifs(0 To 3)
  p.rdifs(0) = diffuser_make((diffscale * b), 0.75)
  p.rdifs(1) = diffuser_make((diffscale * cc), 0.75)
  p.rdifs(2) = diffuser_make((diffscale * dd), 0.625)
  p.rdifs(3) = diffuser_make((diffscale * e), 0.625)


  '/* Tapped delay section */

  p.tapdelay = fixeddelay_make(44000)
  'p.taps = (int *)calloc(FDNORDER, sizeof(int))
  'ReDim p.taps(0 To FDNORDER - 1)
  'p.tapgains = (float *)calloc(FDNORDER, sizeof(float))
  'ReDim p.tapgains(0 To FDNORDER - 1)

  p.taps(0) = 5 + 0.41 * p.largestdelay
  p.taps(1) = 5 + 0.3 * p.largestdelay
  p.taps(2) = 5 + 0.155 * p.largestdelay
  p.taps(3) = 5 + 0# * p.largestdelay

  For i = 0 To FDNORDER - 1
    p.tapgains(i) = (p.alpha ^ p.taps(i))
  Next

  'return(p)
  gverb_new = p
End Function



'
'void gverb_free(ty_gverb * p)
'{
'  int i;
'
'  damper_free(p->inputdamper);
'  for(i = 0; i < FDNORDER; i++) {
'    fixeddelay_free(p->fdndels[i]);
'    damper_free(p->fdndamps[i]);
'    diffuser_free(p->ldifs[i]);
'    diffuser_free(p->rdifs[i]);
'  }
'  free(p->fdndels);
'  free(p->fdngains);
'  free(p->fdnlens);
'  free(p->fdndamps);
'  free(p->d);
'  free(p->u);
'  free(p->f);
'  free(p->ldifs);
'  free(p->rdifs);
'  free(p->taps);
'  free(p->tapgains);
'  fixeddelay_free(p->tapdelay);
'  free(p);
'}
Sub gverb_free(p As ty_gverb)
'{
'  int i
'
'  damper_free(p.inputdamper)
'  for(i = 0  i < FDNORDER  i++) {
'    fixeddelay_free(p.fdndels(i))
'    damper_free(p.fdndamps(i))
'    diffuser_free(p.ldifs(i))
'    diffuser_free(p.rdifs(i))
'  }
'  free(p.fdndels)
'  free(p.fdngains)
'  free(p.fdnlens)
'  free(p.fdndamps)
'  free(p.d)
'  free(p.u)
'  free(p.f)
'  free(p.ldifs)
'  free(p.rdifs)
'  free(p.taps)
'  free(p.tapgains)
'  fixeddelay_free(p.tapdelay)
'  free(p)
End Sub

'
'void gverb_flush(ty_gverb * p)
'{
'  int i;
'
'  damper_flush(p->inputdamper);
'  for(i = 0; i < FDNORDER; i++) {
'    fixeddelay_flush(p->fdndels[i]);
'    damper_flush(p->fdndamps[i]);
'    diffuser_flush(p->ldifs[i]);
'    diffuser_flush(p->rdifs[i]);
'  }
'  memset(p->d, 0, FDNORDER * sizeof(float));
'  memset(p->u, 0, FDNORDER * sizeof(float));
'  memset(p->f, 0, FDNORDER * sizeof(float));
'  fixeddelay_flush(p->tapdelay);
'}
'
Sub gverb_flush(p As ty_gverb)
'{
'  int i
'
'  damper_flush(p.inputdamper)
'  for(i = 0  i < FDNORDER  i++) {
'    fixeddelay_flush(p.fdndels(i))
'    damper_flush(p.fdndamps(i))
'    diffuser_flush(p.ldifs(i))
'    diffuser_flush(p.rdifs(i))
'  }
'  memset(p.d, 0, FDNORDER * sizeof(float))
'  memset(p.u, 0, FDNORDER * sizeof(float))
'  memset(p.f, 0, FDNORDER * sizeof(float))
'  fixeddelay_flush(p.tapdelay)
End Sub

