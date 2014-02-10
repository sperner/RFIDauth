// RegistreUtil.cpp: implementation of the CRegistreUtil class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
//#include "Registre.h"
#include "RegistreUtil.h"
#include <string.h>

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CRegistreUtil::CRegistreUtil()
{

}

CRegistreUtil::~CRegistreUtil()
{

}

int CRegistreUtil::CreerCle (HKEY HK, char * cle) // Cr�� une cl� dans la base de registre
{
	//---------------------------------------------------------------------
	// cr�� une cl� dans la base de registre !
	// Dans la rubrique HK (ex : HKEY_CLASSES_ROOT), et de nom "cl�"
	// Pour cr�er des sous-cl�, on peu directement taper :
	// "cle01\\cle02\\cle03 ...." dans la variable "cl�"
	// La fonction cr�e directement les sous cl�s !
	//---------------------------------------------------------------------

	SECURITY_ATTRIBUTES SecAtt;
	SecAtt.nLength = sizeof (SECURITY_ATTRIBUTES);
	SecAtt.lpSecurityDescriptor = NULL;
	SecAtt.bInheritHandle = TRUE;

	HKEY Retour;
	DWORD Action;


	RegCreateKeyEx(HK,cle,0,"", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, & SecAtt, &Retour, &Action);

	RegCloseKey(Retour);

	return 0;
}

int CRegistreUtil::DetruitCle (HKEY HK, char * cle) // D�truire une cl� dans la base de registre
{
	//------------------------------------------------------------
	// Cette fonction d�truit une cl� dans la base de registre !
	// Vous ne pouvez d�truir qu'une seule cl� a la fois !
	// et cette fonction ne d�truit pas les sous cl�s !
	//------------------------------------------------------------
	RegDeleteKey(HK,cle);

	return 0;

}

int CRegistreUtil::EcrireTexte (HKEY HK, char * cle, char * nom, char * valeur) // Ecrit une valeur Texte dans une cl� d�finie
{
	//------------------------------------------------------------
	// Ecrit une valeur de type chaine de caract�re dans la base de registre
	// "nom" repr�sente le nom de la valeur 
	// "valeur" repr�sente la chaine de caract�re
	//------------------------------------------------------------

	SECURITY_ATTRIBUTES SecAtt;
	SecAtt.nLength = sizeof (SECURITY_ATTRIBUTES);
	SecAtt.lpSecurityDescriptor = NULL;
	SecAtt.bInheritHandle = TRUE;

	HKEY Retour;
	DWORD Action;

	RegCreateKeyEx(HK,cle,0,"", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, & SecAtt, &Retour, &Action);

	
	RegSetValueEx(Retour,nom,0,REG_EXPAND_SZ ,(unsigned char *)valeur,strlen(valeur)+1);


	RegCloseKey(Retour);

	return 0;
}


int CRegistreUtil::EcrireNombre (HKEY HK, char * cle, char * nom, long valeur) // Ecrit une valeur num�rique dans la base de registre
{
	//------------------------------------------------------------
	// Idem EcritTexte, mais la valeur est un nombre
	//------------------------------------------------------------

	SECURITY_ATTRIBUTES SecAtt;
	SecAtt.nLength = sizeof (SECURITY_ATTRIBUTES);
	SecAtt.lpSecurityDescriptor = NULL;
	SecAtt.bInheritHandle = TRUE;

	HKEY Retour;
	DWORD Action;

	RegCreateKeyEx(HK,cle,0,"", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, & SecAtt, &Retour, &Action);

	
	RegSetValueEx(Retour,nom,0,REG_DWORD,(unsigned char *)&valeur,4);


	RegCloseKey(Retour);

	return 0;
}


int CRegistreUtil::LitNombre(HKEY HK, char * cle, char * nom, long *valeur)	// Lit une valeur num�rique dans la base de registre
{
	//------------------------------------------------------------
	// r�cup�re la valeur (num�rique) de la valeur "nom", dans la cl� "cle"
	//------------------------------------------------------------

	HKEY Retour;
	RegOpenKeyEx(HK,cle,0,KEY_ALL_ACCESS,&Retour);

	unsigned long taille = 4; // initialisation obligatoire pour WIN 95- 98
	unsigned long type;
	RegQueryValueEx( Retour,nom,NULL,&type,(unsigned char *)valeur,&taille);

	RegCloseKey(Retour);

	return 0;

}

int CRegistreUtil::LitTexte (HKEY HK, char * cle, char * nom, char * valeur,unsigned long taille) // Lit une valeur alphanum�rique dans la base de registre
{
	//------------------------------------------------------------
	// lit la valeur "nom" dans la cl� "cle"
	// La valeur est une chaine de caract�re
	// La variable TAILLE doit contenir la taille du buffer // IMPERATIF sous win 95-98
	//------------------------------------------------------------

	HKEY Retour;
	RegOpenKeyEx(HK,cle,0,KEY_ALL_ACCESS,&Retour);

	unsigned long type;
	RegQueryValueEx( Retour,nom,NULL,&type,(unsigned char *)valeur,&taille);

	RegCloseKey(Retour);
	return 0;
}

int CRegistreUtil::EnumVal(HKEY HK, char * cle, char **TableauNom, char **TableauVal, int NMax , int MaxCar) // R�cup�res toutes les valeurs d'un cl� de la base de registre
{
	//------------------------------------------------------------
	// Cette fonction lit toutes les valeurs d'une m�me cl� (dans la limite de NMax)
	// La valeur MaxCar d�finie la taille maximum d'une valeur, ainsi que du nom de la valeur
	//
	// Les tableaux, correspondent au donn�e r�cup�r�e dans la base !
	// ATTENTION : LES TABLEAUX DOIVENT ETRES INITIALISES !!!
	// leur taille doit �tre identique, et �gale a NMAX,
	// et chaque chaine du tableau doit �tre �gale a MaxCar
	//------------------------------------------------------------
	HKEY Retour;
	RegOpenKeyEx(HK,cle,0,KEY_ALL_ACCESS,&Retour);
	
	char * NomVal;
	NomVal = new char[MaxCar];
	char * Valeur;
	Valeur = new char[MaxCar];
	unsigned long NNom=MaxCar;
	unsigned long NVal=MaxCar;
	unsigned long Ty=0;
	int n=0;
	long Ret;

	do
	{
		Ret =RegEnumValue( Retour,n,(char *)NomVal,&NNom,0,&Ty,(unsigned char *)Valeur,&NVal );
	
		strcpy(TableauNom[n],NomVal);
		strcpy(TableauVal[n],Valeur);

		n++;
		NNom = MaxCar;
		NVal = MaxCar;
	} while ( (Ret != ERROR_NO_MORE_ITEMS) && (n<NMax) );


	RegCloseKey(Retour);

	return n-1;
}


int CRegistreUtil::EnumCle(HKEY HK, char *cle, char **Tableau,int NMax,int MaxCar) // Liste toutes les sous-cl�s d'une m�me cl�
{
	//------------------------------------------------------------
	// Le tableau doit �tre initialis� a NMAX, et MaxCar
	// NMAX, correspond au nombre maximum de sous cl�s
	// MaxCar est la taille maximum du nom de la sous cl�
	//------------------------------------------------------------

	HKEY Retour;
	RegOpenKeyEx(HK,cle,0,KEY_ALL_ACCESS,&Retour);
	
	char * NomVal;
	NomVal = new char[MaxCar];
	unsigned long NNom=MaxCar;
	int n=0;
	long Ret;

	do
	{
		Ret =RegEnumKeyEx( Retour,n,NomVal,&NNom,NULL,NULL,NULL,NULL );
	
		strcpy(Tableau[n],NomVal);
	
		n++;
		NNom = MaxCar;


	} while ( (Ret != ERROR_NO_MORE_ITEMS) && (n<NMax) );


	RegCloseKey(Retour);
	return n-1;


}