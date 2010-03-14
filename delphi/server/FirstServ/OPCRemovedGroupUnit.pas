unit OPCRemovedGroupUnit;

interface

uses Windows,SysUtils,Dialogs,Classes,Axctrls;

type
  TOPCGroupRemoved = class
  public
   oldGroupName:string;
   oldGroupList:TList;
   removedCount:integer;
   oldServerHandle:longword;
   constructor Create(const aName:string;oldHandle:longword;aList:TList);
  end;

implementation

constructor TOPCGroupRemoved.Create(const aName:string;oldHandle:longword;aList:TList);
begin
 oldGroupName:=aName;
 oldServerHandle:=oldHandle;
 oldGroupList:=aList;
 removedCount:=0;
end;

end.
