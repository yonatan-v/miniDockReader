# miniDockReader

This is a small and minimal modern C++ library for reading DOCX documents. The library was developed especially for those who want to add support for importing DOCX documents with standard formatting to their software and do not want to get involved with heavy libraries.

The library consists of two files:
**miniDocxReader.h** - the header file
**miniDocxReader.cpp** - the source file

Also, the library uses two third-party projects:

**tinyxml2** - for reading XML files embedded in the DOCX file (https://github.com/leethomason/tinyxml2)

**miniz** - for opening the DOCX file compressed in ZIP format (https://github.com/tfussell/miniz-cpp)
