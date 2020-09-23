Note added 2001-06-04: This latest update allows you to use the Enhanced Cryptographic Service provider, so you can use longer keys.  You can now select your key length, too.  Thanks to Tina Tark for providing me with the correct Enhanced CSP name.  You can NOT use the Enhanced CSP without the high-encryption Service Pack for NT4, or an update for Win2K, both available from: http://www.microsoft.com/technet/security/crypload.asp

----------------------------------------------------------------------
Another note, 2001-06-04:  This is a test project, designed to show you how to use the CryptoAPI.  It has a poor user interface design (I threw it together for your benefit) so don't use this as an example of good UI design.  Read up on good UI design if you are making professional apps.  You must have it!
I also know that not all encrypted characters display in the text boxes on the single-key and key-pair forms.  You just can't display control characters (tab, enter, backspace, etc.), and in most cases the VB text boxes don't try...rightfully so.  You will need to add code to convert the output to hex or something like it if you must display encrypted text.  I just displayed it as garbage so you get positive feedback from the app that the text was encrypted.
----------------------------------------------------------------------

This wrapper class serves as a wrapper for the Microsoft CryptoAPI (Base CSP) and the Zlib compression dll.  This document includes descriptions of both sets of functionality and their use within a program.  Keep in mind when you write your own code that if you plan on compressing AND encrypting a file you will achieve best results if you compress it first and then encrypt it.  This is because if you do it backwards and encrypt first, the file is effectively randomized and it will be much harder for the compression algorithms to find repeating data to compress.


**********************************************************************
Encryption/Signature Use
**********************************************************************

***Single key encryption:

-instantiate the object

-use the SessionStart method

-use any of the single key methods (they do not have a _KeyPair extension)

-EncryptFile stores the ValueSALT property for you in the file.  If you use EncryptString or EncryptByteArray you must handle storing ValueSALT and restoring ValueSALT before calling DecryptString or DecryptByteArray.  ValueSALT is not encrypted, and can be stored along with the encrypted data without compromising security at all.

-if you have many things to encrypt or decrypt do them all at once

-use the SessionEnd method when finished

-set the object = nothing


***Key pair encryption:

-instantiate the object

-use the SessionStart method

-generate new key pair with Generate_KeyPair (and then usually export it with Export_KeyPair for use later; get them from ValuePublicKey and ValuePublicPrivateKey) OR import previously saved key(s) by setting ValuePublicKey/ValuePublicPrivateKey and calling Import_KeyPair.  ValuePublicPrivateKey should never be released to anyone else.  The private key is encrypted like PGP does.  See http://www.pgp.com/products/freeware/default.asp for the best (free) commercial-grade file & email encryption program around.

-use any of the key pair methods (they have the _KeyPair extension)

-EncryptFile_KeyPair stores the ValueSessionKey property for you in the file.  If you use EncryptString_KeyPair or EncryptByteArray_KeyPair you must handle storing ValueSessionKey and restoring ValueSessionKey before calling DecryptString_KeyPair or DecryptByteArray_KeyPair.  ValueSessionKey is the random session key used to encrypt the data.  That session key is encrypted using the key pair.  The encrypted session key is what is stored in ValueSessionKey, and this can be stored along with the encrypted data.  

-use the SessionEnd method when finished

-set the object = nothing


***Signature/Validation:

-instantiate the object

-use the SessionStart method

-use any of the Sign???_KeyPair or Validate???_KeyPair methods

-SignFile_KeyPair stores the ValueSignature property for you in the file.  If you use SignString or SignByteArray you must handle storing ValueSignature and restoring ValueSignature before calling ValidateString or ValidateByteArray.  ValueSignature is not encrypted, and can be stored along with the data it signs without compromising security at all.

-if you have many things to sign or validate do them all at once

-use the SessionEnd method when finished

-set the object = nothing
**********************************************************************
Encryption/Signature functionality:

This wrapper class encrypts/decrypts/signs/validates a string, byte array, or file (using the CryptoAPI included with Windows) using a single user-entered key, OR key pair (RSA) encryption.  This is not another "you can't break my new encryption" snippet written by a 12-year-old (I guarantee the NSA can break every one on this site with only a calculator, in 5 minutes).  The CryptoAPI uses business/government grade encryption/signature methods designed for use in business applications.  If you want encryption ignore everything but the CryptoAPI.  Period.

This code was adapted from code that Fredrik Qvarfort adapted (he demonstrated 7 REAL encryption algorithms using VB code plus the CryptoAPI).  He showed how to use the CryptoAPI for single-key encryption.  I wanted to focus on the API for use within my company so I improved on that one object.  My major addition to the single key encryption is the ability to generate a random SALT and set the SALT.  If you don't know what a SALT is you must read all of Microsoft's site on using the CryptoAPI:
http://msdn.microsoft.com/library/psdk/crypto/portalapi_3351.htm?RLD=290
They explain how cryptography and the CryptoAPI works (sort of; I had to discover some things on my own when it came to programming key pair encryption).  It is beyond code notation or this README file to explain it.  Complicated at first, but by following this code you see what they mean.

In this latest version I have finally figured out how to do public/private (a.k.a. key pair or RSA) encryption and signing.  It took me over two solid weeks, 8 hours each day, to do it.  Figuring out APIs in VB is fun, but what a pain in the ass, too!  I hereby post this effort on planetsourcecode because I couldn't find CryptoAPI key pair encryption in VB or C code anywhere on the net (P.S. after finishing this code I did find one example in VB; it was very complicated and I figured my example was far easier to understand for someone just starting out with the CryptoAPI, believe it or not).  They don't even include full examples in C on MSDN.  I am sure others can use this, so feel free to use it.

I added two new methods (SessionStart and SessionEnd) to open and close the crypt context handle.  This is for performance reasons.  If you need to encrypt/sign multiple items it is much faster to get a handle, do all the encryption/signatures, and then close the handle when finished.

This class only uses the Base Cryptographic Service Provider.  There are several others that are widely available (via NT Service Packs, etc.) but I have not explored their use yet.  The greatest perks to using other CSPs are the ability to generate longer keys, and use stronger encryption algorithms, all for better security.  I leave that for another week.

**********************************************************************
Compression Use:
**********************************************************************

-instantiate the object

-use any method

-CompressFile stores the length of the original file for you.  That value is necessary for decompression.  If you use CompressString or CompressByteArray you must store the ValueDecompressedSize property and restore that property before calling DecompressString or DecompressByteArray.

-set the object = nothing

**********************************************************************
Compression functionality:

This wrapper class and example program demonstrates the proper use of the Zlib compression dll.  Unbelievably, Windows does not provide adequate string/file compression in the API (only decompression of files created with the compress.exe or compact.exe utilities that come with Windows, and their compression is weak by today's standards).  No string or array compression support...and having to call an exe to compress a file is far from perfect.

To fill this void a group wrote the Zlib.dll utility (included; their web site address is http://www.info-zip.org/pub/infozip/zlib so you can get the latest version).  The standards committee that created the PNG picture format used the Zlib compression code as the standard for PNG.  So if you have used PNG you have indirectly used Zlib.  It provides fast and compact encryption for byte arrays.  This wrapper class extends it to strings and full files, too.  Two other postings on www.planetsourcecode.com deal with Zlib.  One is barely functional and the other only provides the coverted C header files (but is otherwise excellent with many utilities and modules to do other things -- check out "Kira" posted by The_Lung).  I found this code on the Zlib web site, and converted it from an ocx to a regular class module to conserve resources.  I also moved all the code dealing with file compression into the class where it belonged.

Doug Gaede
dgaede@home.com
December 11, 2000
