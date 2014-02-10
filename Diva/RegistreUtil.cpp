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

int CRegistreUtil::CreerCle (HKEY HK, char * cle) // Créé une clé dans la base de registre
{
	//---------------------------------------------------------------------
	// créé une clé dans la base de registre !
	// Dans la rubrique HK (ex : HKEY_CLASSES_ROOT), et de nom "clé"
	// Pour créer des sous-clé, on peu directement taper :
	// "cle01\\cle02\\cle03 ...." dans la variable "clé"
	// La fonction crée directement les sous clés !
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

int CRegistreUtil::DetruitCle (HKEY HK, char * cle) // Détruire une clé dans la base de registre
{
	//------------------------------------------------------------
	// Cette fonction détruit une clé dans la base de registre !
	// Vous ne pouvez détruir qu'une seule clé a la fois !
	// et cette fonction ne détruit pas les sous clés !
	//------------------------------------------------------------
	RegDeleteKey(HK,cle);

	return 0;

}

int CRegistreUtil::EcrireTexte (HKEY HK, char * cle, char * nom, char * valeur) // Ecrit une valeur Texte dans une clé définie
{
	//------------------------------------------------------------
	// Ecrit une valeur de type chaine de caractère dans la base de registre
	// "nom" représente le nom de la valeur 
	// "valeur" représente la chaine de caractère
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


int CRegistreUtil::EcrireNombre (HKEY HK, char * cle, char * nom, long valeur) // Ecrit une valeur numérique dans la base de registre
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


int CRegistreUtil::LitNombre(HKEY HK, char * cle, char * nom, long *valeur)	// Lit une valeur numérique dans la base de registre
{
	//------------------------------------------------------------
	// récupère la valeur (numérique) de la valeur "nom", dans la clé "cle"
	//------------------------------------------------------------

	HKEY Retour;
	RegOpenKeyEx(HK,cle,0,KEY_ALL_ACCESS,&Retour);

	unsigned long taille = 4; // initialisation obligatoire pour WIN 95- 98
	unsigned long type;
	RegQueryValueEx( Retour,nom,NULL,&type,(unsigned char *)valeur,&taille);

	RegCloseKey(Retour);

	return 0;

}

int CRegistreUtil::LitTexte (HKEY HK, char * cle, char * nom, char * valeur,unsigned long taille) // Lit une valeur alphanumérique dans la base de registre
{
	//------------------------------------------------------------
	// lit la valeur "nom" dans la clé "cle"
	// La valeur est une chaine de caractère
	// La variable TAILLE doit contenir la taille du buffer // IMPERATIF sous win 95-98
	//------------------------------------------------------------

	HKEY Retour;
	RegOpenKeyEx(HK,cle,0,KEY_ALL_ACCESS,&Retour);

	unsigned long type;
	RegQueryValueEx( Retour,nom,NULL,&type,(unsigned char *)valeur,&taille);

	RegCloseKey(Retour);
	return 0;
}

int CRegistreUtil::EnumVal(HKEY HK, char * cle, char **TableauNom, char **TableauVal, int NMax , int MaxCar) // Récupères toutes les valeurs d'un clé de la base de registre
{
	//------------------------------------------------------------
	// Cette fonction lit toutes les valeurs d'une même clé (dans la limite de NMax)
	// La valeur MaxCar définie la taille maximum d'une valeur, ainsi que du nom de la valeur
	//
	// Les tableaux, correspondent au donnée récupérée dans la base !
	// ATTENTION : LES TABLEAUX DOIVENT ETRES INITIALISES !!!
	// leur taille doit être identique, et égale a NMAX,
	// et chaque chaine du tableau doit être égale a MaxCar
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


int CRegistreUtil::EnumCle(HKEY HK, char *cle, char **Tableau,int NMax,int MaxCar) // Liste toutes les sous-clés d'une même clé
{
	//------------------------------------------------------------
	// Le tableau doit être initialisé a NMAX, et MaxCar
	// NMAX, correspond au nombre maximum de sous clés
	// MaxCar est la taille maximum du nom de la sous clé
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