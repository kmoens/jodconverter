JODConverter for Cipal IT Solutions nv
======================================
A fork of the JODConverter library which specific modifications for Cipal 
IT Solutions nv.

The following patches are applied:
* A check on the path length of the instance profile directory on Windows. 
If it exceeds 159 characters, OpenOffice fails to start properly.
* Remove of the subdirectory for the instance profile directory. We make our 
work directory unique outside of the library.
* Remove of the Sigar library, it is unstable in production use.

JODConverter
============

JODConverter (for Java OpenDocument Converter) automates document conversions
using LibreOffice or OpenOffice.org.

I started this project back in 2003, but I am no longer maintaining it. I moved
the code here at GitHub in the hope that a well-maintained fork will emerge.

See the [Google Code](http://code.google.com/p/jodconverter/) project for more
info.

