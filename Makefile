

CC=cl.exe /nologo
CXX=cl.exe /nologo
TARGET=aPSL
VSDIR=C:\Program Files\Microsoft Visual Studio 10.0\VC
MSSDK=C:\Program Files\Microsoft SDKs\Windows\v7.0A
CXXFLAGS=/EHsc \
		 /I"$(VSDIR)\include" \
		 /I"$(MSSDK)\Include" \
		 /IPSL
LDFLAGS=/DLL \
		/LIBPATH:"$(VSDIR)\Lib" \
		/LIBPATH:"$(MSSDK)\Lib"
LIBS=Advapi32.lib comsuppw.lib
REGSVR=regsvr32.exe
FILTER=iconv -f SJIS -t UTF-8 | tee build.log

all:
	$(MAKE) install 2>&1 |  $(FILTER)

PSL:
	git clone git://github.com/saitoha/PSL 
#	git clone git@github.com:saitoha/PSL 

install: $(TARGET).dll
	$(REGSVR) /s $(TARGET).dll

uninstall: $(TARGET).dll
	$(REGSVR) /s /u $(TARGET).dll
	
$(TARGET).dll: $(TARGET).cpp $(TARGET).def Makefile PSL
	$(CXX) $(CXXFLAGS) \
		$(TARGET).cpp \
		/link $(LDFLAGS) \
		/DEF:$(TARGET).def \
		/OUT:$@ \
		$(LIBS) 

clean:
	$(RM) *.obj *.dll *.exp *.lib *log

