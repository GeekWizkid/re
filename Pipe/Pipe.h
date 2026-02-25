// The following ifdef block is the standard way of creating macros which make exporting 
// from a DLL simpler. All files within this DLL are compiled with the PIPE_EXPORTS
// symbol defined on the command line. this symbol should not be defined on any project
// that uses this DLL. This way any other project whose source files include this file see 
// PIPE_API functions as being imported from a DLL, whereas this DLL sees symbols
// defined with this macro as being exported.
#ifdef PIPE_EXPORTS
#define PIPE_API __declspec(dllexport)
#else
#define PIPE_API __declspec(dllimport)
#endif

// This class is exported from the Pipe.dll
class PIPE_API CPipe {
public:
	CPipe(void);
	// TODO: add your methods here.
};

extern PIPE_API int nPipe;

PIPE_API int fnPipe(void);
