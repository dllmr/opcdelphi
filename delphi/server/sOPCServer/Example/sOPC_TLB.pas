unit sOPC_TLB;

// ************************************************************************ //
// WARNUNG                                                                    
// -------                                                                    
// Die in dieser Datei deklarierten Typen wurden aus Daten einer Typbibliothek
// generiert. Wenn diese Typbibliothek explizit oder indirekt (über eine     
// andere Typbibliothek) reimportiert wird oder wenn die Anweisung            
// 'Aktualisieren' im Typbibliotheks-Editor während des Bearbeitens der     
// Typbibliothek aktiviert ist, wird der Inhalt dieser Datei neu generiert und 
// alle manuell vorgenommenen Änderungen gehen verloren.                           
// ************************************************************************ //

// PASTLWTR : 1.2
// Datei generiert am 08.03.2006 08:41:38 aus der unten beschriebenen Typbibliothek.

// ************************************************************************  //
// Typbib: C:\sProject\sOPC\Delphi\sOPCServer\Example\DataAccess20Demo.tlb (1)
// LIBID: {118921D1-0703-11D5-962A-00A024AEBA44}
// LCID: 0
// Hilfedatei: 
// Hilfe-String: sOPC Demo OPC DA2 Server Library
// DepndLst: 
//   (1) v2.0 stdole, (C:\WINDOWS\system32\stdole2.tlb)
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit muß ohne Typüberprüfung für Zeiger compiliert werden. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
interface

uses Windows, ActiveX, Classes, Graphics, StdVCL, Variants;
  

// *********************************************************************//
// In dieser Typbibliothek deklarierte GUIDS . Es werden folgende         
// Präfixe verwendet:                                                     
//   Typbibliotheken     : LIBID_xxxx                                     
//   CoClasses           : CLASS_xxxx                                     
//   DISPInterfaces      : DIID_xxxx                                      
//   Nicht-DISP-Schnittstellen: IID_xxxx                                       
// *********************************************************************//
const
  // Haupt- und Nebenversionen der Typbibliothek
  sOPCMajorVersion = 1;
  sOPCMinorVersion = 0;

  LIBID_sOPC: TGUID = '{118921D1-0703-11D5-962A-00A024AEBA44}';

  IID_IOPCDataAccess20: TGUID = '{118921D2-0703-11D5-962A-00A024AEBA44}';
  IID_IOPCGroup: TGUID = '{118921D4-0703-11D5-962A-00A024AEBA44}';
  CLASS_OPCGroup: TGUID = '{118921D6-0703-11D5-962A-00A024AEBA44}';
  CLASS_OPCDataAccess20: TGUID = '{118921D8-0703-11D5-962A-00A024AEBA44}';
type

// *********************************************************************//
// Forward-Deklaration von in der Typbibliothek definierten Typen         
// *********************************************************************//
  IOPCDataAccess20 = interface;
  IOPCDataAccess20Disp = dispinterface;
  IOPCGroup = interface;

// *********************************************************************//
// Deklaration von in der Typbibliothek definierten CoClasses             
// (HINWEIS: Hier wird jede CoClass zu ihrer Standardschnittstelle        
// zugewiesen)                                                            
// *********************************************************************//
  OPCGroup = IOPCGroup;
  OPCDataAccess20 = IOPCDataAccess20;


// *********************************************************************//
// Schnittstelle: IOPCDataAccess20
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {118921D2-0703-11D5-962A-00A024AEBA44}
// *********************************************************************//
  IOPCDataAccess20 = interface(IDispatch)
    ['{118921D2-0703-11D5-962A-00A024AEBA44}']
  end;

// *********************************************************************//
// DispIntf:  IOPCDataAccess20Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {118921D2-0703-11D5-962A-00A024AEBA44}
// *********************************************************************//
  IOPCDataAccess20Disp = dispinterface
    ['{118921D2-0703-11D5-962A-00A024AEBA44}']
  end;

// *********************************************************************//
// Schnittstelle: IOPCGroup
// Flags:     (0)
// GUID:      {118921D4-0703-11D5-962A-00A024AEBA44}
// *********************************************************************//
  IOPCGroup = interface(IUnknown)
    ['{118921D4-0703-11D5-962A-00A024AEBA44}']
  end;

// *********************************************************************//
// Die Klasse CoOPCGroup stellt die Methoden Create und CreateRemote zur      
// Verfügung, um Instanzen der Standardschnittstelle IOPCGroup, dargestellt von
// CoClass OPCGroup, zu erzeugen. Diese Funktionen können                     
// von einem Client verwendet werden, der die CoClasses automatisieren    
// möchte, die von dieser Typbibliothek dargestellt werden.               
// *********************************************************************//
  CoOPCGroup = class
    class function Create: IOPCGroup;
    class function CreateRemote(const MachineName: string): IOPCGroup;
  end;

// *********************************************************************//
// Die Klasse CoOPCDataAccess20 stellt die Methoden Create und CreateRemote zur      
// Verfügung, um Instanzen der Standardschnittstelle IOPCDataAccess20, dargestellt von
// CoClass OPCDataAccess20, zu erzeugen. Diese Funktionen können                     
// von einem Client verwendet werden, der die CoClasses automatisieren    
// möchte, die von dieser Typbibliothek dargestellt werden.               
// *********************************************************************//
  CoOPCDataAccess20 = class
    class function Create: IOPCDataAccess20;
    class function CreateRemote(const MachineName: string): IOPCDataAccess20;
  end;

implementation

uses ComObj;

class function CoOPCGroup.Create: IOPCGroup;
begin
  Result := CreateComObject(CLASS_OPCGroup) as IOPCGroup;
end;

class function CoOPCGroup.CreateRemote(const MachineName: string): IOPCGroup;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_OPCGroup) as IOPCGroup;
end;

class function CoOPCDataAccess20.Create: IOPCDataAccess20;
begin
  Result := CreateComObject(CLASS_OPCDataAccess20) as IOPCDataAccess20;
end;

class function CoOPCDataAccess20.CreateRemote(const MachineName: string): IOPCDataAccess20;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_OPCDataAccess20) as IOPCDataAccess20;
end;

end.
