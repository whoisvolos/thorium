'_________________________________________________________________________________________
' mt19337ar.vb version 2.3 (for Visual Basic .NET), 2008-12-02:
'
' This file implements the Mersenne Twister (MT), MT19937ar,
' pseudo-random number generator (PRNG) as a Visual Basic .NET
' class.
'
' Add this file (only) to your VB project to use the Mersenne Twister.
'
' This version (2.3) is unchanged functionally from version 2.2.  Only comments
' were changed from "Visual Basic 2005" to "Visual Basic .NET", reflecting
' the fact that the code was and is compatible with VB 2008 .NET and VB 2005 .NET.
'
' I marked almost every major change I made with "Visual Basic .NET" and
' bounded the changes with low lines: '_____________ ....  The primary exceptions
' are expressions that used uAdd(), uMult(), uDiv() and uDiv2(); they
' now use native Visual Basic .NET operators but are unmarked.
'
' To see a list of PRNGs herein, search for "- FUNCTIONS" (no quotes).
' To see instructions for using the PRNGs, search for "- USAGE" (no quotes).
' To see performance results, search for "- ON PERFORMANCE" (no quotes).
'
' Class MTRandom was implemented in three steps:
'
'	1) C-program for MT19937, with initialization improved 2002/1/26.  Coded by Takuji
'	   Nishimura and Makoto Matsumoto.
'
'	2) Translation to VBA was made and tested by Pablo Mariano Ronchi (2005-Sep-12).
'
'	3) Translation from Visual Basic for Applications (VBA) to Visual Basic .NET
'	   was made and tested by Ron Charlton (2008-09-23). He also translated
'	   genrand_intMax() from C++ code by Richard J. Wagner and Magnus Jonsson.
'
'		Ron Charlton
'		9002 Balcor Circle
'		Knoxville, TN 37923-2301 USA
'		Phone:		1-865-694-0800
'		Email:		charltoncr @ comcast.net (remove spaces)
'		Home Page:	http://home.comcast.net/~charltoncr/#MersenneTwister
'		Date:		2008-09-23
'
' Comments from all three steps are included.
'
' Bug reports, improvements, corrections, complaints or kudos are welcomed.
'
' The VBA version was obtained via a link at
' http://www.math.sci.hiroshima-u.ac.jp/~m-mat/MT/emt.html on 2008-07-24.
' The original MT C-code can also be found at the same web address.
'
' For cryptographically secure random number generation, search
' Visual Basic .NET Help for "RandomNumberGenerator Class". Or
' see http://www.math.sci.hiroshima-u.ac.jp/~m-mat/MT/efaq.html.
'
' Changes made for the Visual Basic .NET version:
'	o Changed all "Long" declarations to "UInteger" per VB .NET Help.
'	o Changed all "CLng()" calls to "CUInt()" calls.
'	o Made the code a VB Class and added New()s.
'	o Moved mti and mt() variables into a class so they can be saved
'	  to a file and loaded later to restore a generator's state.
'	o Added init_random(), saveState() and loadState().
'	o Added genrand_intMax(), genrand_intRange() and genrand_int64().
'	o Added parentheses to function calls.
'	o Replaced uAdd(), uMult(), uDiv() and uDiv2() with native VB operators.
'	o Moved Main() to a separate file, "Main.vb", for ease of use.
'	o Copied SimplePerformanceTest to a separate file and revised it.
'	o Files "Demo.vb" and "MakeEntropyTestData.vb" contain entirely new
'	  Visual Basic .NET code.
'
' I compared the Visual Basic .NET Main()'s output file,
' "mt19937arVBTest.out", to file "mt19937ar.out" from the original MT
' authors' web site. The files are identical.  (See "USAGE" below
' for details.)
'
' I also modified the original mt19937ar.c code and this VB code to produce
' 100,000,000 random integers and 100,000,000 random reals.  The C & VB
' outputs are identical.
'
' Note: In Visual Basic .NET --
'	Data type Integer  has a range of -2^31 inclusive to 2^31-1 inclusive.
'	Data type UInteger has a range of   0   inclusive to 2^32-1 inclusive.
'	Data type Long     has a range of -2^63 inclusive to 2^63-1 inclusive.
'	Data type ULong    has a range of   0   inclusive to 2^64-1 inclusive.
'__________________________________________________________________________________________



'This [was] the Visual Basic for Applications (VBA) version of the  MT19937ar,
'or   "MERSENNE TWISTER"   algorithm for pseudo random number generation,
'with initialization improved, by  MAKOTO MATSUMOTO  and  TAKUJI NISHIMURA,
'of  2002/1/26.
'This translation to VBA was made and tested by Pablo Mariano Ronchi (2005-Sep-12)

'Note 1: VBA is the Visual Basic language used in MS Access, MS Excel and, in general,
'        in MS Office, and is called simply "Visual Basic" or VBA, hereinafter.
'Note 2: This same code compiles in Visual Basic (VB) without modifications.
' [Note 2 applied to an earlier version of VB.  It was not true for VB .NET. (Ron C.)]


'Please read the comments about this VBA version that follow the ones below, by
'the authors of the "MERSENNE TWISTER" algorithm.




'/*
'   A C-program for MT19937, with initialization improved 2002/1/26.
'   Coded by Takuji Nishimura and Makoto Matsumoto.
'
'   Before using, initialize the state by using init_genrand(seed)
'   or init_by_array(init_key, key_length).
'
'   Copyright (C) 1997 - 2002, Makoto Matsumoto and Takuji Nishimura,
'   All rights reserved.
'
'   Redistribution and use in source and binary forms, with or without
'   modification, are permitted provided that the following conditions
'   are met:
'
'     1. Redistributions of source code must retain the above copyright
'        notice, this list of conditions and the following disclaimer.
'
'     2. Redistributions in binary form must reproduce the above copyright
'        notice, this list of conditions and the following disclaimer in the
'        documentation and/or other materials provided with the distribution.
'
'     3. The names of its contributors may not be used to endorse or promote
'        products derived from this software without specific prior written
'        permission.
'
'   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
'   "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
'   LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
'   A PARTICULAR PURPOSE ARE DISCLAIMED.  IN NO EVENT SHALL THE COPYRIGHT OWNER OR
'   CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
'   EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
'   PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR
'   PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
'   LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
'   NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
'   SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'
'
'   Any feedback is very welcome.
'   http://www.math.sci.hiroshima-u.ac.jp/~m-mat/MT/emt.html
'   email: m-mat @ math.sci.hiroshima-u.ac.jp (remove space)
'*/

' It would be nice to CC: rjwagner@writeme.com, Cokus@math.washington.edu,
' pmronchi@yahoo.com.ar and charltoncr @ comcast.net when you write.


'COMMENTS ABOUT THIS VISUAL BASIC FOR APPLICATIONS (VBA) VERSION OF
'THE  "Mersenne Twister"  ALGORITHM:
'
'- All the statements made by the authors of the original algorithm and implementation,
'  present in the original C code and copied above, including but not limiting to those
'  regarding the license, the usage without any warranties, and the conditions for
'  distribution, apply to this Visual Basic program for testing that algorithm.
'
'
'- If you use this Visual Basic version, just for the record please send me an email to:
'
'  pmronchi@yahoo.com.ar
'
'  with an appropriate reference to your work, the city and country where you reside,
'  and the word "MT19937ar" in the subject. Thanks.
'
'
'- FUNCTIONS AND PROCEDURES IMPLEMENTED:
'
'       Function                    Returns values in the range:
'       -------------------------   -------------------------------------------------------
'       genrand_int32()             [0, 4294967295]  (0 to 2^32-1)
'
'       genrand_int31()             [0, 2147483647]  (0 to 2^31-1)
'
'       genrand_real1()             [0.0, 1.0]   (both 0.0 and 1.0 included)
'
'       genrand_real2()             [0.0, 1.0) = [0.0, 0.9999999997672...]
'                                                (0.0 included, 1.0 excluded)
'
'       genrand_real3()             (0.0, 1.0) = [0.0000000001164..., 0.9999999998836...]
'                                                (both 0.0 and 1.0 excluded)
'
'       genrand_res53()             [0.0,~1.0] = [0.0, 1.00000000721774...]
'                                                (0.0 included, ~1.0 included)
'
'       The following ADDITIONAL functions
'       ARE NOT PRESENT IN THE ORIGINAL C CODE:
'
'       NOTE: the limits shown below, marked with (*), are valid if gap==5.0e-13
'
'       genrand_int32SignedInt()   [-2147483648, 2147483647]   (-2^31 to 2^31-1)
'
'       genrand_real2b()           [0.0, 1.0)=[0, 1-(2*gap)] =[0.0, 0.9999999999990] (*)
'                                             (0.0 included, 1.0 excluded)
'
'       genrand_real2c()           (0.0, 1.0]=[0+(2*gap),1.0]=[1.0e-12, 1.0] (*)
'                                             (0.0 excluded, 1.0 included)
'
'       genrand_real3b()           (0.0, 1.0)=[0+gap, 1-gap] =[5.0e-13, 0.9999999999995] (*)
'                                             (both 0.0 and 1.0 excluded)
'
'
'       (See the "Acknowledgements" section for the following functions)
'
'       genrand_real4b()           [-1.0,1.0]=[-1.0, 1.0]
'                                             (-1.0 included, 1.0 included)
'
'       genrand_real5b()           (-1.0,1.0)=[-1.0+(2*gap), 1.0-(2*gap)]=
'                                                    [-0.9999999999990,0.9999999999990] (*)
'                                             (-1.0 excluded, 1.0 excluded)
'_________________________________________________________________________________________
'		These three functions were added for the Visual Basic .NET version. 
'
'		genrand_intMax(upper)			[0,upper] for upper < 2^32	(0 to 4294967295 but <= upper)
'
'		genrand_intRange(lower,upper)	[lower,upper] for 0 <= lower <= upper <= 2^32-1
'													(0 <= lower <= upper <= 4294967295)
'
'		genrand_int64()					[0,18446744073709551615]	(0 to 2^64-1)
'_________________________________________________________________________________________
'
'
'       Procedure                 Arguments
'       ------------------------  ---------------------------------------------------------
'       init_genrand(seed)        any seed included in [0,4294967295]
'       init_by_array(array)      array has elements of type UInteger;
'                                 the array must have at least one element
'
'		[Visual Basic .NET]
'       init_random(True|False)   True:  reseed VB Random from the system clock,
'										 then reseed MTRandom from VB Random
'                                 False: use the next value from VB Random to
'										 reseed MTRandom
'		saveState(fileName)		  any valid file name
'		loadState(fileName)		  name of any file saved earlier with saveState()
'
'
'- USAGE:
'_________________________________________________________________________________________
' In your Visual Basic .NET application:
'
'	1) Add this file to your Visual Basic .NET project.
'
'	2) Create one or more instances of the MTRandom class.  The pseudo-random sequence
'	   for each instance is initialized based on the arguments.
'	   Make certain each instance stays in scope for the duration of its use so
'	   it won't be re-initialized inadvertently.  Five different methods of
'	   initialization are:
'
'			a) Dim r As New MTRandom()
'					Initialize with seed == 5489 (the original MT
'					authors' default seed).
'
'			b) Dim r As New MTRandom(19456)
'					Initialize with seed==19456.  Any unsigned integer seed
'					in the range [0,4294967295] is acceptable.
'
'			c) Dim init() As UInteger = {&H123, &H234, &H345, &H456}
'			   Dim r As New MTRandom(init)
'					Initialize with an array of unsigned integers (as in the
'					original MT authors' version).  The array must have
'					at least one element. 624 elements are desirable.  More
'					or fewer elements are acceptable.  The element values
'					should be random.
'
'			d) Dim r As New MTRandom(True)
'					Seed VB Random from the system clock, and then
'					seed MTRandom with the next value from VB Random.
'					True is required but ignored.
'
'			e) Dim r As New MTRandom("filename.ext")
'					Initialize with a generator state saved earlier with
'					MTRandom.saveState("filename.ext")
'
'	3) Call any of the genrand_X() functions listed above.  You can re-
'      initialize an MTRandom instance at any point by calling init_genrand(),
'	   init_by_array() or init_random().  You can save or load the PRNG state
'	   at any point by calling saveState() or loadState().
'
'	4) Catch exceptions as desired.  The MTRandom exception classes are at the
'	   end of this file.
'
' To test genrand_int32(), genrand_real2(), init_genrand() and init_by_array()
' in the Visual Basic .NET version of MT:
'
'	1) Create a Visual Basic .NET console application Project consisting of files
'			mt19937ar.vb
'			Demo.vb
'			SimplePerformanceTest.vb
'			MakeEntropyTestData.vb
'			Main.vb
'
'	2) Build and run the project.
'
'	3) Compare the file "mt19937arVBTest.out" with the original MT authors'
'	   test output file ("mt19937ar.out", found by following
'	   links at http://www.math.sci.hiroshima-u.ac.jp/~m-mat/MT/emt.html).  
'	   They should be identical.
'_________________________________________________________________________________________
'
'
'- ON THE NEED OF THE Mersenne Twister ALGORITHM:
'
'  If you are serious about (pseudo) randomness in your Visual Basic project, then
'  you MUST use this Visual Basic version of the Mersenne Twister algorithm.
'  DO NOT RELY on the randomize() and rnd() routines provided by Visual Basic!
'
'
'- ON PERFORMANCE:
'_________________________________________________________________________________________
'  SimplePerformanceTest.vb test results for MTRandom (Visual Basic .NET)
'
'	On a Dell Precision 390 workstation with 2.66 GHz Intel Core2 Duo processor,
'	3.5 GB RAM, Visual Basic .NET compiler, and generating 100,000,000 random numbers
'	with each function (running the Release version at a Cmd.exe prompt):
'
'									Seconds		Nanoseconds		Calls Per
'	Function						Run Time	Per Call		Second
'	--------------------------		--------	-----------		----------
'	genrand_int32()					 1.656		  16.56			60,390,000
'	genrand_real1()					 2.031		  20.31			49,240,000
'	genrand_res53()		 			 4.250		  42.50			23,530,000
'	genrand_intMax(1000)			 2.343		  23.43			42,680,000
'	genrand_intMax(&H5FFFFFFF)		 3.359		  33.59			29,770,000
'	genrand_intRange(10,20)			 3.813		  38.13			26,230,000
'	genrand_int64()					 3.516		  35.16			28,440,000
'_________________________________________________________________________________________
'
'
'
'  - If you are using Visual Basic, then performance is hardly an issue, and very probably
'    these routines will not be a bottleneck. Anyway, I made my best to optimize the code
'    for speed without compromising its clarity and readability, and while strictly
'    following the original C code.
'    If you need better performance, consider using the C or PHP version of the
'    Mersenne Twister algorithm, and MySQL for database management, for instance.
'
'  - This new version (September 2005), is around 9% faster than the previous one
'    published on-line (April 2005).
'
'  - Performance tests: For my own testing procedures I have developed a full test.
'    I would like to make the test package publicly available, but as of this writing
'    I have not asked Mr.Makoto Matsumoto yet. If he agrees, you will probably
'    find it close to the place were you found this code.
'
'  - Just to give you an idea of the difference in performance between C and VBA, one
'    of the parts of my test is a loop that only calls genrand_int32() repeatedly,
'    generating 100 million (1e08) pseudo-random numbers.
'    The times in my old Pentium MMX, 120Mhz, for this part of the test are:
'         6008 seconds (1h40m) for VBA, and 56 seconds for C (a relation of 107:1).
'
'  - A simple VBA code to test the performance is given below (SimplePerformanceTest()).
'    The C code provided in the main() function by the authors of the Mersenne Twister
'    algorithm could be easily adapted to match the result of a suitable modified
'    SimplePerformanceTest():
'
'    Public Sub SimplePerformanceTest()
'    Const kMaxNr As Integer = 1000000    'CHANGE THIS VALUE AS NEEDED
'    Dim ii As Integer, tmp As Double, sec1 As Double, sec2 As Double
'    Open "mt19937arVBtest.txt" For Output As #1    'open the output file
'
'    sec1 = Timer
'    For ii = 1 To kMaxNr: tmp = genrand_int32(): Next 'use any of the functions
'    sec2 = Timer: tmp = sec2 - sec1
'
'    Print #1, "Elapsed time in seconds for generating "; Trim(kMaxNr); " numbers: "; _
'              Fix(tmp) & "." & Format(Fix((tmp - Fix(tmp) + 0.005) * 100), "00")
'    Close #1    'close the output file
'    End Sub 'SimplePerformanceTest
'
'
'- [VBA] DIFFERENCES WITH THE ORIGINAL C FUNCTIONS AND SOURCE FILE:
'
'    [The section on why genrand_int32() returned Double was removed.  genrand_int32()
'	 now returns UInteger.]
'
'  - There are minor changes (addition and use of variable mtb) to produce the same result
'    in Visual Basic as in C, when a genrand_X() function is called without a previous
'    call to one of the init_X procedures. I apologize for using this not very elegant
'    patch, but there is no simpler way to simulate the use of the "static" word in C,
'    given that the VBA's "static" word does not behave in a similar way.
'    Another minor change, for similar reasons, is the declaration and initialization of
'    the mag01() array.
'
'  - I added many constants, for clarity in some cases, and also for speed in others,
'    because some operations could be faster if defined as a multiplication instead of
'    a division, in some processors.
'
'  - I made small changes in the main() function, in order to easily change the number
'    of output values in the listings.
'
'
'- IS THIS VBA CODE DEPENDABLE?
'
'  Well, I think so. I have tested the code by generating 101026001 pseudo-random numbers
'  using all of the functions, printed 389 of them, and compared the result with a similar
'  test I wrote in C using the original code of the Mersenne Twister algorithm. Both
'  outputs were exactly the same, except for the timings and some variable headings.
'
'  Besides, I kept the original C code commented out in this source file, in order
'  to easily check and understand the translation to Visual Basic. So, you will find
'  a block of one or a few lines of Visual Basic code following the corresponding
'  -commented out- original C block.
'  This way you can easily see that I have strictly followed the original C code, except
'  for the minor differences explained in the above section, and verify that the VBA code
'  is an almost exact "copy" of the original algorithm.
'  This fact provides another level of confidence to the end result.
'
'  No bugs or problems were reported since the publication on-line of the previous
'  version (April 2005).
'
'
'- ACKNOWLEDGEMENTS:
'
'  I want to thank Mr.MAKOTO MATSUMOTO and Mr.TAKUJI NISHIMURA for creating and
'  generously sharing their excellent algorithm.
'
'  I also want to thank my friends Alejandra María Ribichich, Mariana Francisco Mera,
'  and Víctor Fernando Torres, for their inspiration and support.
'
'  My friend Claudio Pacciarini clarified me some of the differences between VB and VBA.
'
'  Mr.Mutsuo Saito, assistant to professor Matsumoto, patiently e-mailed with me and
'  performed the necessary tasks for the previous version (April 2005) to appear on-line.
'
'  Mr. Kenneth C. Ives (USA) and Mr. David Grundy (Hong Kong) were the first ones
'  aknowledging the use of this VBA code. Thanks for your feedback!
'
'  Mr. Kenneth C. Ives also sent me some code and the idea in which I based
'  genrand_real4b() and genrand_real5b()
'
'
'  Pablo Mariano Ronchi
'  Buenos Aires, Argentina
'
'
'
'End of comments for the Visual Basic for Applications version


Option Strict On
Option Explicit On

'_________________________________________________________________________________________
'  For Visual Basic .NET
Imports System.Xml
Imports System.Xml.Serialization
Imports System.IO
<Serializable()>
Public Class MTRandom

    ' a class (within the MTRandom class) to hold the MT generator state
    <Serializable()> Public Class MTState
        Public mti As Integer
        Public mt(Nuplim) As UInteger

        Sub New()
            Me.mti = Nplus1
        End Sub
    End Class

    ' Here lies the state of a generator so it can be saved to an XML file
    ' and loaded later.
    Dim state As New MTState()

    ' a VB Random Class PRNG used in init_random() to seed MTRandom based
    ' on the system clock
    Private rng As Random = Nothing
    '_________________________________________________________________________________________


    '#include <stdio.h>
    '
    '/* Period parameters */
    '#define N 624
    '#define M 397
    '#define MATRIX_A 0x9908b0dfUL   /* constant vector a */
    '#define UPPER_MASK 0x80000000UL /* most significant w-r bits */
    '#define LOWER_MASK 0x7fffffffUL /* least significant r bits */
    Private Const N As Integer = 624
    Private Const M As Integer = 397
    Private Const MATRIX_A As UInteger = &H9908B0DFUI       '/* constant vector a */
    Private Const UPPER_MASK As UInteger = &H80000000UI '/* most significant w-r bits */
    Private Const LOWER_MASK As UInteger = &H7FFFFFFFUI '/* least significant r bits */

    'To avoid unnecesary operations while using the Visual Basic interpreter:
    Private Const kDiffMN As Integer = M - N
    Private Const Nuplim As Integer = N - 1
    Private Const Muplim As Integer = M - 1
    Private Const Nplus1 As Integer = N + 1
    Private Const NuplimLess1 As Integer = Nuplim - 1
    Private Const NuplimLessM As Integer = Nuplim - M

    'static unsigned long mt[N];  /* the array for the state vector */
    'static int mti=N+1;          /* mti==N+1 means mt[N] is not initialized */
    ' The following two VBA lines are replaced by code at the beginning of
    ' class MTRandom.
    ''Dim mt(0 To Nuplim) As Integer	 '/* the array for the state vector */
    ''Dim mti As Integer = Nplus1

    'In the C original version the following array, mag01(), is declared within
    'the function genrand_int32(). In VBA I had to declare it global for performance
    'considerations, and because there is no way in VBA to emulate the use of the word
    '"static" in C:
    '
    'static unsigned long mag01[2]={0x0UL, MATRIX_A};
    '/* mag01[x] = x * MATRIX_A  for x=0,1 */
    Dim mag01(2) As UInteger

    'Other constants defined to be used in this Visual Basic version:

    'Powers of 2: k2_X means 2^X
    Private Const k2_8 As Integer = 256
    Private Const k2_16 As Integer = 65536
    Private Const k2_24 As Integer = 16777216

    Private Const k2_31 As Double = 2147483648.0#     '2^31   ==  2147483648 == 80000000
    Private Const k2_32 As Double = 2.0# * k2_31      '2^32   ==  4294967296 == 0
    Private Const k2_32b As Double = k2_32 - 1.0#     '2^32-1 ==  4294967295 == FFFFFFFF == -1

    'The following constant has its value defined by the authors of the
    'Mersenne Twister algorithm
    Private Const kDefaultSeed As UInteger = 5489


    'The following constant, is used within genrand_real1(), which returns values in [0,1]
    Private Const kMT_1 As Double = 1.0# / k2_32b

    'The following constant, is used within genrand_real2(), which returns values in [0,1)
    Private Const kMT_2 As Double = 1.0# / k2_32

    'The following constant, is used within genrand_real3(), which returns values in (0,1)
    Private Const kMT_3 As Double = kMT_2

    'The following constant, used within genrand_res53(), is needed, because the Visual
    'Basic interpreter cannot read real LITERALS with the the same precision as a C compiler,
    'and so ends up truncating the least significant decimal digit(s), a '2' in this case.
    'The original factor used in the C code is: 9007199254740992.0
    Private Const kMT_res53 As Double = 1.0# / (9.00719925474099E+15 + 2.0#)    'add lost digit '2'


    'The following constants, are used within the ADDITIONAL functions genrand_real2b() and
    'genrand_real3b(), equivalent to genrand_real() and genrand_real3(), but that return
    'evenly distributed values in the ranges [0, 1-kMT_Gap] and [0+kMT_Gap, 1-kMT_Gap],
    'respectively. A similar statement is valid also for genrand_real2c(), genrand_real4b()
    'and genrand_real5b(). See the section "Functions and procedures implemented" above,
    'for more details.
    '
    'If you want to change the value of kMT_Gap, it is suggested to do it so that:
    '   5e-15 <= kMT_Gap <= 5e-2

    Private Const kMT_Gap As Double = 0.0000000000005       '5.0E-13
    Private Const kMT_Gap2 As Double = 2.0# * kMT_Gap         '1.0E-12
    Private Const kMT_GapInterval As Double = 1.0# - kMT_Gap2 '0.9999999999990

    Private Const kMT_2b As Double = kMT_GapInterval / k2_32b
    Private Const kMT_2c As Double = kMT_2b
    Private Const kMT_3b As Double = kMT_2b
    Private Const kMT_4b As Double = 2.0# / k2_32b
    Private Const kMT_5b As Double = (2.0# * kMT_GapInterval) / k2_32b   '1.999999999998/k2_32b





    '_________________________________________________________________________________________
    ' For Visual Basic .NET
    ' ----------[MTRandom Contructors]-----------

    ' initialize the PRNG with the default seed
    Public Sub New()
        '  init_genrand(5489UL); /* a default initial seed is used */
        Me.init_genrand(kDefaultSeed)
    End Sub

    ' initialize the MTRandom PRNG with seed
    Public Sub New(ByVal seed As UInteger)
        Me.init_genrand(seed)
    End Sub

    ' initialize the MTRandom PRNG with an unsigned integer array
    Public Sub New(ByRef array() As UInteger)
        Me.init_by_array(array)
    End Sub

    ' Initialize the MTRandom PRNG with a pseudo-random seed.  Variable dummy
    ' is used only to distinguish this overload from the others.
    Public Sub New(ByVal dummy As Boolean)
        Me.init_random(True)
    End Sub

    ' initialize the MTRandom PRNG with an XML file created by MTRandom.saveState()
    Public Sub New(ByVal fileName As String)
        Me.loadState(fileName)
    End Sub
    ' ----------[End of MTRandom Contructors]-----------


    ' Initialize the MTRandom PRNG with a pseudo-random seed.
    ' reSeedFromClock: True - reseed the INITIALIZER RNG from the system clock
    '				   False - use the next RN from the INITIALIZER RNG
    Public Sub init_random(ByVal reSeedFromClock As Boolean)

        ' If this is the first call of init_random() or user asks for reseed
        ' of rng from the system clock.
        ' (Must check for rng Is Nothing because user might re-initialize
        ' the MTRandom instance when it was first initialized another way.)
        If rng Is Nothing Or reSeedFromClock Then
            ' seed rng from the system clock by making a new instance
            rng = New Random()
        End If

        ' initialize MTRandom with a pseudo-random Integer from rng
        Me.init_genrand(CUInt(CLng(rng.Next(Int32.MinValue, Int32.MaxValue)) - CLng(Int32.MinValue)))
    End Sub

    ' Save the PRNG state to a file as XML
    Public Sub saveState(ByVal fileName As String)
        Try
            Dim serializer As New XmlSerializer(GetType(MTState))
            Dim fs As New FileStream(fileName, FileMode.Create, FileAccess.Write)

            serializer.Serialize(fs, state)
            fs.Close()
        Catch ex As Exception
            Throw New MTRandomSaveStateException(ex.Message, Environment.StackTrace)
        End Try
    End Sub

    ' Load the PRNG state from a file created by MTRandom.saveState()
    Public Sub loadState(ByVal fileName As String)
        Try
            Dim serializer As New XmlSerializer(GetType(MTState))
            Dim fs As New FileStream(fileName, FileMode.Open, FileAccess.Read)

            state = DirectCast(serializer.Deserialize(fs), MTState)
            fs.Close()
        Catch ex As Exception
            Throw New MTRandomLoadStateException(ex.Message, Environment.StackTrace)
        End Try

    End Sub
    '_________________________________________________________________________________________




    Public Sub init_genrand(ByVal seed As UInteger)      'void init_genrand(unsigned long s)
        '/* initializes mt[N] with a seed */
        'mt[0]= s & 0xffffffffUL;
        'for (mti=1; mti<N; mti++) {
        '    mt[mti] =
        '    (1812433253UL * (mt[mti-1] ^ (mt[mti-1] >> 30)) + mti);
        '    /* See Knuth TAOCP Vol2. 3rd Ed. P.106 for multiplier. */
        '    /* In the previous versions, MSBs of the seed affect   */
        '    /* only MSBs of the array mt[].                        */
        '    /* 2002/01/09 modified by Makoto Matsumoto             */
        '    mt[mti] &= 0xffffffffUL;
        '    /* for >32 bit machines */

        Dim tt As UInteger

        state.mt(0) = (seed And &HFFFFFFFFUI)
        For state.mti = 1 To Nuplim
            'original expression, rearranged in one line:
            'mt[mti] = (1812433253UL * (mt[mti-1] ^ (mt[mti-1] >> 30)) + mti);

            tt = state.mt(state.mti - 1)
            state.mt(state.mti) = CUInt((1812433253UL * CULng(tt Xor (tt >> 30)) + CULng(state.mti)) And &HFFFFFFFFUL)
            ' The next statement is incorporated into the previous statement.
            'state.mt(state.mti) = state.mt(state.mti) And &HFFFFFFFFUI	 '/* for >32 bit machines */
        Next

        mag01(0) = 0 : mag01(1) = MATRIX_A
    End Sub     'init_genrand

    Public Sub init_by_array(ByRef init_key() As UInteger)
        'void init_by_array(unsigned long init_key[], int key_length)

        '/* initialize by an array with array-length */
        '/* init_key is the array for initializing keys */
        '/* key_length is its length */
        '/* slight change for C++, 2004/2/26 */

        'int i, j, k;
        Dim i As Integer, j As Integer, k As Integer
        Dim key_length As Integer = init_key.GetUpperBound(0) + 1
        Dim tt As UInteger


        'init_genrand(19650218UL);
        'i=1; j=0;
        'k = (N>key_length ? N : key_length);
        init_genrand(19650218UI)
        i = 1 : j = 0
        k = CInt(IIf((N > key_length), N, key_length))


        'for (; k; k--) {
        For k = k To 1 Step -1  'while k<>0, that is: while k>0
            'original expression, rearranged in one line:
            'mt[i] = (mt[i] ^ ((mt[i-1] ^ (mt[i-1] >> 30)) * 1664525UL)) + init_key[j] + j;

            tt = state.mt(i - 1)
            state.mt(i) = CUInt((CULng((state.mt(i) Xor ((tt Xor (tt >> 30))) * 1664525UL)) + CULng(init_key(j)) + CULng(j)) And &HFFFFFFFFUL)

            'mt[i] &= 0xffffffffUL;          /* for WORDSIZE > 32 machines */
            ''unnecesary, due to previous statement

            'i++; j++;
            'if (i>=N) { mt[0] = mt[N-1]; i=1; }
            'if (j>=key_length) j=0;
            i = i + 1 : j = j + 1
            If i >= N Then state.mt(0) = state.mt(Nuplim) : i = 1
            If j >= key_length Then j = 0
        Next


        'for (k=N-1; k; k--) {
        For k = Nuplim To 1 Step -1
            'original expression, rearranged in one line:
            'mt[i] = (mt[i] ^ ((mt[i-1] ^ (mt[i-1] >> 30)) * 1566083941UL)) - i;  /* non linear */

            tt = state.mt(i - 1)
            state.mt(i) = CUInt(((CULng(state.mt(i)) Xor CULng((tt Xor (tt >> 30)) * 1566083941UL)) - CULng(i)) And &HFFFFFFFFUL)

            'mt[i] &= 0xffffffffUL;          /* for WORDSIZE > 32 machines */
            ''unnecesary, due to previous statement

            'i++;
            'if (i>=N) { mt[0] = mt[N-1]; i=1; }
            i = i + 1
            If i >= N Then state.mt(0) = state.mt(Nuplim) : i = 1
        Next


        'mt[0] = 0x80000000UL;   /* MSB is 1; assuring non-zero initial array */
        state.mt(0) = &H80000000UI      '/* MSB is 1; assuring non-zero initial array */
    End Sub     'init_by_array


    ' genrand_int32SignedInt() is only for compatiblity with the VBA version. It does not
    ' serve the same purpose as in the VBA version.
    Public Function genrand_int32SignedInt() As Integer
        Dim tmp As Long = genrand_int32()

        If tmp > Int32.MaxValue Then
            tmp -= UInt32.MaxValue + 1
        End If

        Return CInt(tmp)
    End Function

    Public Function genrand_int32() As UInteger 'unsigned long genrand_int32(void)
        '/* generates a random number on [0,0xffffffff]-interval */

        'unsigned long y;
        Dim y As UInteger

        'The below lines were replaced by another approach. See section "On performance" for details:
        'static unsigned long mag01[2]={0x0UL, MATRIX_A};
        '/* mag01[x] = x * MATRIX_A  for x=0,1 */

        If (state.mti >= N) Then    '{ /* generate N words at one time */
            'int kk;
            Dim kk As Integer

            'if (mti == N+1)   /* if sgenrand() has not been called, */
            '  init_genrand(5489UL); /* a default initial seed is used */
            If state.mti = Nplus1 Then init_genrand(kDefaultSeed)

            'for (kk=0;kk<N-M;kk++) {
            '    y = (mt[kk]&UPPER_MASK)|(mt[kk+1]&LOWER_MASK);
            '    mt[kk] = mt[kk+M] ^ (y >> 1) ^ mag01[y & 0x1UL];
            '}
            For kk = 0 To (NuplimLessM)
                y = (state.mt(kk) And UPPER_MASK) Or (state.mt(kk + 1) And LOWER_MASK)
                state.mt(kk) = state.mt(kk + M) Xor (y >> 1) Xor mag01(CInt(y And 1UI))
            Next

            'for (;kk<N-1;kk++) {
            '    y = (mt[kk]&UPPER_MASK)|(mt[kk+1]&LOWER_MASK);
            '    mt[kk] = mt[kk+(M-N)] ^ (y >> 1) ^ mag01[y & 0x1UL];
            '}
            For kk = kk To NuplimLess1
                y = (state.mt(kk) And UPPER_MASK) Or (state.mt(kk + 1) And LOWER_MASK)
                state.mt(kk) = state.mt(kk + (M - N)) Xor (y >> 1) Xor mag01(CInt(y And 1UI))
            Next

            'y = (mt[N-1]&UPPER_MASK)|(mt[0]&LOWER_MASK);
            'mt[N-1] = mt[M-1] ^ (y >> 1) ^ mag01[y & 0x1UL];
            y = (state.mt(Nuplim) And UPPER_MASK) Or (state.mt(0) And LOWER_MASK)
            state.mt(N - 1) = state.mt(M - 1) Xor (y >> 1) Xor mag01(CInt(y And 1UI))
            'mti = 0;
            state.mti = 0
        End If


        y = state.mt(state.mti) : state.mti = state.mti + 1
        '/* Tempering */
        'y ^= (y >> 11);
        y = y Xor (y >> 11)

        'y ^= (y << 7) & 0x9d2c5680UL;
        y = y Xor (y << 7) And &H9D2C5680UI

        'y ^= (y << 15) & 0xefc60000UL;
        y = y Xor (y << 15) And &HEFC60000UI

        'y ^= (y >> 18);
        'return y;
        Return y Xor (y >> 18)
    End Function    'genrand_int32

    Public Function genrand_int31() As Integer    'long genrand_int31(void)
        '/* generates a random number on [0,0x7fffffff]-interval */
        'return (long)(genrand_int32()>>1);
        Return CInt(genrand_int32() >> 1)
    End Function    'genrand_int31

    Public Function genrand_real1() As Double   'double genrand_real1(void)
        '/* generates a random number on [0,1]-real-interval */
        'return genrand_int32()*(1.0/4294967295.0);     '/* divided by 2^32-1 */
        Return genrand_int32() * kMT_1
    End Function    'genrand_real1

    Public Function genrand_real2() As Double   'double genrand_real2(void)
        '/* generates a random number on [0,1)-real-interval */
        'return genrand_int32()*(1.0/4294967296.0);     '/* divided by 2^32 */
        Return genrand_int32() * kMT_2
    End Function    'genrand_real2

    Public Function genrand_real3() As Double   'double genrand_real3(void)
        '/* generates a random number on (0,1)-real-interval */
        'return (((double)genrand_int32()) + 0.5)*(1.0/4294967296.0);   '/* divided by 2^32 */
        Return (CDbl(genrand_int32()) + 0.5) * kMT_3
    End Function    'genrand_real3

    Public Function genrand_res53() As Double   'double genrand_res53(void)
        '/* generates a random number on [0,1) with 53-bit resolution*/
        'unsigned long a=genrand_int32()>>5, b=genrand_int32()>>6;
        'return(a*67108864.0+b)*(1.0/9007199254740992.0);
        Return kMT_res53 * ((genrand_int32() >> 5) * 67108864.0# + (genrand_int32() >> 6))
    End Function    'genrand_res53

    '/* These (PREVIOUS) real versions are due to Isaku Wada, 2002/01/09 added */


    'The following functions are present only in the Visual Basic version, not in the
    'C version. See more comments in the definition of the constants used as factors:

    Public Function genrand_real2b() As Double
        'Returns results in the range [0,1) == [0, 1-kMT_Gap2]
        'Its lowest value is : 0.0
        'Its highest value is: 0.9999999999990
        Return genrand_int32() * kMT_2b
    End Function    'genrand_real2b

    Public Function genrand_real2c() As Double
        'Returns results in the range (0,1] == [0+kMT_Gap2, 1.0]
        'Its lowest value is : 0.0000000000010  (1E-12)
        'Its highest value is: 1.0
        Return kMT_Gap2 + (genrand_int32() * kMT_2c)    '==kMT_Gap2+genrand_real2b()
    End Function    'genrand_real2c

    Public Function genrand_real3b() As Double   'double genrand_real3(void)
        'Returns results in the range (0,1) == [0+kMT_Gap, 1-kMT_Gap]
        'Its lowest value is : 0.0000000000005  (5E-13)
        'Its highest value is: 0.9999999999995
        Return kMT_Gap + (genrand_int32() * kMT_3b)
    End Function    'genrand_real3b

    'Mr. Kenneth C. Ives sent me some code and the idea in which I based genrand_real4b() and
    'genrand_real5b(). Added on 2005-Sep-12:

    Public Function genrand_real4b() As Double
        'Returns results in the range [-1,1] == [-1.0, 1.0]
        'Its lowest value is : -1.0
        'Its highest value is: 1.0
        Return (genrand_int32() * kMT_4b) - 1.0#
    End Function    'genrand_real4b

    Public Function genrand_real5b() As Double
        'Returns results in the range (-1,1) == [-kMT_GapInterval, kMT_GapInterval]
        'Its lowest value is : -0.9999999999990
        'Its highest value is: 0.9999999999990
        Return kMT_Gap2 + ((genrand_int32() * kMT_5b) - 1.0#)
    End Function    'genrand_real5b

    '__________________________________________________________________________________________
    ' The following functions were added by Ron Charlton 2008-09-23 for Visual Basic .NET.

    Public Function genrand_intMax(ByVal N As UInteger) As UInteger
        ' Returns a UInteger in [0,n] for 0 <= n < 2^32
        ' Its lowest value is : 0
        ' Its highest value is: 4294967295 but <= N

        ' Translated by Ron Charlton from C++ file 'MersenneTwister.h' where it is named
        ' MTRand::randInt(const uint32& n), and has the following comments:
        '-----
        ' Mersenne Twister random number generator -- a C++ class MTRand
        ' Based on code by Makoto Matsumoto, Takuji Nishimura, and Shawn Cokus
        ' Richard J. Wagner  v1.0  15 May 2003  rjwagner@writeme.com

        ' Copyright (C) 1997 - 2002, Makoto Matsumoto and Takuji Nishimura,
        ' Copyright (C) 2000 - 2003, Richard J. Wagner
        ' All rights reserved.                          
        '-----
        ' MersenneTwister.h can be found at
        ' http://www-personal.umich.edu/~wagnerr/MersenneTwister.html.

        ' Find which bits are used in N
        ' Optimized by Magnus Jonsson (magnus@smartelectronix.com)
        Dim used As UInteger = N
        used = used Or (used >> 1)
        used = used Or (used >> 2)
        used = used Or (used >> 4)
        used = used Or (used >> 8)
        used = used Or (used >> 16)

        ' Draw numbers until one is found in [0,n]
        Dim i As UInteger
        Do
            i = genrand_int32() And used    ' toss unused bits to shorten search
        Loop While i > N

        Return i
    End Function

    Public Function genrand_intRange(ByVal lower As UInteger, ByVal upper As UInteger) As UInteger
        ' Generate a pseudo-random integer between lower inclusive and upper inclusive for
        ' 0 <= lower <= upper <= 4294967295.
        ' Returns a UInteger in the range [lower,upper].
        ' Its lowest value is : 0 but >= lower
        ' Its highest value is: 4294967295 but <= upper
        '
        ' Written by Ron Charlton, 2008-09-23.
        Return lower + genrand_intMax(upper - lower)
    End Function

    Public Function genrand_int64() As ULong
        ' Returns an unsigned long in the range [0,2^64-1]
        ' Its lowest value is : 0
        ' Its highest value is: 18446744073709551615
        '
        ' Written by Ron Charlton, 2008-09-23.

        Return CULng(genrand_int32()) Or (CULng(genrand_int32()) << 32)
    End Function
End Class



' ----------[EXCEPTION CLASSES]----------

' The "unable to save the PRNG's state" exception for MTRandom.
' Properties:
'	Message		- an error description
'	StackTrace	- a stack trace
Class MTRandomSaveStateException
	Inherits Exception

	Private msg As String
	Private stkTrace As String	' stack trace

	Public Sub New(ByVal MyBaseMessage As String, ByVal stackTrace As String)
		Const NL As String = ControlChars.NewLine

		msg = "Failed to save generator state in MTRandom.saveState:" & NL & _
		 MyBaseMessage

		stkTrace = stackTrace
	End Sub

	Public Overrides ReadOnly Property StackTrace() As String
		Get
			Return stkTrace
		End Get
	End Property

	Public Overrides ReadOnly Property Message() As String
		Get
			Return msg
		End Get
	End Property

End Class



' The "unable to load the PRNG's state" exception for MTRandom.
' Properties:
'	Message		- an error description
'	StackTrace	- a stack trace
Class MTRandomLoadStateException
	Inherits Exception

	Private msg As String
	Private stkTrace As String	' stack trace

	Public Sub New(ByVal MyBaseMessage As String, ByVal stackTrace As String)
		Const NL As String = ControlChars.NewLine

		msg = "Failed to load generator state in MTRandom.loadState:" & NL & _
		 MyBaseMessage

		stkTrace = stackTrace
	End Sub

	Public Overrides ReadOnly Property StackTrace() As String
		Get
			Return stkTrace
		End Get
	End Property

	Public Overrides ReadOnly Property Message() As String
		Get
			Return msg
		End Get
	End Property

End Class
'__________________________________________________________________________________________
