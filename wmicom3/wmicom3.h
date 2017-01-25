#pragma once

#ifdef WMICOM3_EXPORTS  
#define WMICOM3_API __declspec(dllexport)   
#endif  

extern "C"
{
	WMICOM3_API int wmicom(char *);
}
