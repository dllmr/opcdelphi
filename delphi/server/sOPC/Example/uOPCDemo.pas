unit uOPCDemo;

interface

uses
  Variants, SysUtils, ActiveX, OPCDA, uOPCNode, uOPC;

// OPC functions
procedure OnInitAddressSpace;
procedure OnRead(OPCNode: sOPCNode; Path: string; DataType: TVarType);
procedure OnWrite(OPCNode: sOPCNode; Value: variant; Path: string; DataType: TVarType);

// Value functions
procedure IncrementValues;
procedure ResetValues;

var
  // actual values (Tags)
  Values: array[101 .. 404] of integer;

implementation

procedure IncrementValues;
var
  i: integer;
begin
  for i := Low(Values) to High(Values) do Values[i] := Values[i] + 1;
end;

procedure ResetValues;
var
  i: integer;
begin
  for i := Low(Values) to High(Values) do Values[i] := 0;
end;

procedure OnInitAddressSpace;
// initialize address space
begin
  with OPC do begin
    // add leaf (tag)
    AddLeaf(0, 101, 'String',  VT_BSTR,          OPC_READABLE);
    AddLeaf(0, 102, 'Integer', VT_I4,            OPC_READABLE + OPC_WRITEABLE);
    AddLeaf(0, 103, 'Double',  VT_R8,            OPC_READABLE + OPC_WRITEABLE);
    AddLeaf(0, 104, 'Array',   VT_R8 + VT_ARRAY, OPC_READABLE);

    // add branch
    AddBranch(0, 200, 'Branch 1');
      AddLeaf(200, 201, 'String',  VT_BSTR,          OPC_READABLE);
      AddLeaf(200, 202, 'Integer', VT_I4,            OPC_READABLE + OPC_WRITEABLE);
      AddLeaf(200, 203, 'Double',  VT_R8,            OPC_READABLE + OPC_WRITEABLE);
      AddLeaf(200, 204, 'Array',   VT_R8 + VT_ARRAY, OPC_READABLE);

      AddBranch(200, 300, 'Branch 1.1');
        AddLeaf(300, 301, 'String',  VT_BSTR,          OPC_READABLE);
        AddLeaf(300, 302, 'Integer', VT_I4,            OPC_READABLE + OPC_WRITEABLE);
        AddLeaf(300, 303, 'Double',  VT_R8,            OPC_READABLE + OPC_WRITEABLE);
        AddLeaf(300, 304, 'Array',   VT_R8 + VT_ARRAY, OPC_READABLE);

    AddBranch(0, 400, 'Branch 2');
      AddLeaf(400, 401, 'String',  VT_BSTR,          OPC_READABLE);
      AddLeaf(400, 402, 'Integer', VT_I4,            OPC_READABLE + OPC_WRITEABLE);
      AddLeaf(400, 403, 'Double',  VT_R8,            OPC_READABLE + OPC_WRITEABLE);
      AddLeaf(400, 404, 'Array',   VT_R8 + VT_ARRAY, OPC_READABLE);
  end;
end;

procedure OnRead(OPCNode: sOPCNode; Path: string; DataType: TVarType);
type
  DoubleArray = array[0 .. 65535] of double;
  PDoubleArray = ^DoubleArray;
var
  pda: PDoubleArray;
  i: integer;
begin
  case OPCNode.Ident of

    // string
    101, 201, 301, 401: begin
      VarClear(OPCNode.CurrentValue);
      OPCNode.LastUpdate := Now;
      OPCNode.Quality := OPC_QUALITY_GOOD;
      OPCNode.CurrentValue := Format('Value %d is %d', [OPCNode.Ident, Values[OPCNode.Ident]]);
    end;

    // integer
    102, 202, 302, 402: begin
      { When the Tag represent a status information or when a Tag is changing
        unregularly and you want to generate a data callback only when the value
        has changed, then set the OPCNode fields only when the cached value
        (CurrentValue) is not equal to the actual value.
        The Result is, that the timestamp on the client will be updated only
        when the Tag has changed!
      }
      // On the first OnRead call, CurrentValue is empty!
      if not VarIsEmpty(OPCNode.CurrentValue) then begin
        // Tag has not changed -> no data callback -> exit
        if Values[OPCNode.Ident] = OPCNode.CurrentValue then exit;
      end;
      VarClear(OPCNode.CurrentValue);
      OPCNode.LastUpdate := Now;
      OPCNode.Quality := OPC_QUALITY_GOOD;
      OPCNode.CurrentValue := Values[OPCNode.Ident];
    end;

    // double
    103, 203, 303, 403: begin
      VarClear(OPCNode.CurrentValue);
      OPCNode.LastUpdate := Now;
      OPCNode.Quality := OPC_QUALITY_GOOD;
      OPCNode.CurrentValue := Values[OPCNode.Ident] / 100;
    end;

    // array of double
    104, 204, 304, 404: begin
      VarClear(OPCNode.CurrentValue);
      OPCNode.LastUpdate := Now;
      OPCNode.Quality := OPC_QUALITY_GOOD;
      OPCNode.CurrentValue := VarArrayCreate([0, 31], VT_R8);
      pda := VarArrayLock(OPCNode.CurrentValue);
      for i := 0 to 31 do pda[i] := Values[OPCNode.Ident] / 100;
      VarArrayUnlock(OPCNode.CurrentValue);
    end;

  end;
end;

procedure OnWrite(OPCNode: sOPCNode; Value: variant; Path: string; DataType: TVarType);
begin
  OPCNode.LastUpdate := Now;
  Values[OPCNode.Ident] := Value;
  OPCNode.CurrentValue := Value;
end;

initialization
  FillChar(Values, sizeof(Values), 0);
end.

