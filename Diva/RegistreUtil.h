// RegistreUtil.h: interface for the CRegistreUtil class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_REGISTREUTIL_H__3DE9C219_2936_439A_B8B3_8EE99A390573__INCLUDED_)
#define AFX_REGISTREUTIL_H__3DE9C219_2936_439A_B8B3_8EE99A390573__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


//	Les valeurs HKEY sont définies dans les librairies MFC ! Elle sont donc utilisables
//	directement ! En fait, la valeur correspond au nom réel dans la base de registre :
//	HKEY_CLASSES_ROOT
//	HKEY_CURRENT_USER
//	HKEY_LOCAL_MACHINE
//	HKEY_USERS
//	HKEY_CURRENT_CONFIG




class CRegistreUtil  
{
public:
	CRegistreUtil();
	virtual ~CRegistreUtil();

	static int CreerCle(HKEY HK, char * cle); 
	static int DetruitCle(HKEY HK, char * cle);
	static int EcrireTexte(HKEY HK, char * cle, char * nom, char * valeur);
	static int EcrireNombre(HKEY HK, char * cle, char * nom, long valeur);
	static int LitTexte(HKEY HK, char * cle, char * nom, char * valeur, unsigned long taille);
	static int LitNombre(HKEY HK, char * cle, char * nom, long *valeur);
	static int EnumVal(HKEY HK, char * cle, char **TableauNom, char ** TableauVal, int NMax ,int MaxCar);
	static int EnumCle(HKEY HK, char *cle, char **Tableau,int NMax,int MaxCar);
};

#endif // !defined(AFX_REGISTREUTIL_H__3DE9C219_2936_439A_B8B3_8EE99A390573__INCLUDED_)
