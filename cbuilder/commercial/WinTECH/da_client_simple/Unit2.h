//---------------------------------------------------------------------------
#ifndef Unit2H
#define Unit2H
//---------------------------------------------------------------------------
/////////////////////////////////////////////////////////////////////////////
//
// CItemDef defines the local instance for each defined item
//
//////////////////////////////////////////////////////////////////////////////
class TItemDef : public TObject
{
public:
	__fastcall TItemDef(void);
	__fastcall ~TItemDef(void);
        

        String     Name;		// Item Name
        String     Path;		// Path Name
	HANDLE		Handle;			// Handle returned from the Server
	DWORD		AccessRights;	// Access Rights     "    "    "
	FILETIME	TimeStamp;		// Current TimeStamp, Quality & Value
	DWORD		Quality;		// Updated from the Server
	VARIANT		Value;
};

#endif
 