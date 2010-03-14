
/////////////////////////////////////////////////////////////////////////////////////////////
//
//	Callback Definitions
//
//	Used by the DLL to pass control back to the Server Application
//	Each callback must be explicitly enabled by calling one of the following
//	exported API functions:
//					EnableWriteNotification (WRITENOTIFYPROC lpCallback);
//					EnableUnknownItemNotification (UNKNOWNITEMPROC lpCallback);
//					EnableItemRemovalNotification (ITEMREMOVEDPROC lpCallback);
//					EnableDisconnectNotification (DISCONNECTPROC lpCallback);
//					EnableEventMsgs (EVENTMSGPROC lpCallback);
//					EnableRateNotification (RATECHANGEPROC lpCallback);
//					EnableDeviceRead (DEVICEREADPROC lpCallback);
//		
//
//	WRITENOTIFYPROC
//		Signals the Application that an OPC Client has requested a write to a 
//		tag.  HANDLE value represents the tag, VARIANT* defines the new data,
//		and DWORD* allows the application to return either S_OK or an ERROR result
//	UNKNOWNITEMPROC
//		Signals the Application that a Client has requested a tag that has
//		not been previously defined, (or has been suspended).  The two character
//		pointer arguments represent the OPC Path Name and Item Name of the requested 
//		tag.  This callback allows the Server Application to operate with dynamic
//		tags that are created and maintained only when subscribed to by a connected
//		Client.
//	ITEMREMOVEDPROC
//		Works in conjuction with the above UNKNOWNITEMPROC.  This callback signals the
//		Application that the last client subscription for a given item has ended.  The
//		HANDLE of the tag item is returned as well as the Path and Item Name.  The
//		Server Application may choose to remove or suspend the tag when no clients are
//		subscribing to the data.  
//	DISCONNECTPROC
//		Notifies the Server Application whenever a client disconnects.  The DWORD argument
//		defines the number of client connections remaining.  This callback may be used
//		to shutdown the server when the last client disconnects.
//	EVENTMSGPROC
//		Allows the Application to receive event messages from the WtOPCsvr.dll for
//		tracing OPC Client Interface calls.  Primarily used for debugging purposes.
//
//	RATECHANGEPROC
//		Notifies the Application of the fastest update rate requested by a client for an item.
//		May be used to selectively change the server side data refresh rate to optimize
//		operation for different client connections.
//
//	DEVICEREADPROC
//		Callback into Application whenever a client requests a SyncIO Read
//		with dwSource set to OPC_DS_DEVICE.  Argument list includes the item
//		handle with pointers to the data, quality, and timestamp to be supplied
//		by the application
/////////////////////////////////////////////////////////////////////////////////////////////
#ifdef STRICT
typedef VOID (CALLBACK* WRITENOTIFYPROC)(HANDLE, VARIANT*, DWORD*);
typedef VOID (CALLBACK* UNKNOWNITEMPROC)(CHAR*,CHAR*);
typedef VOID (CALLBACK* ITEMREMOVEDPROC)(HANDLE,CHAR*,CHAR*);
typedef VOID (CALLBACK* DISCONNECTPROC)(DWORD);
typedef VOID (CALLBACK* EVENTMSGPROC)(CHAR*);
typedef VOID (CALLBACK* RATECHANGEPROC)(HANDLE, DWORD);
typedef VOID (CALLBACK* DEVICEREADPROC)(HANDLE, VARIANT*, WORD*, FILETIME*);
#else /* !STRICT */
typedef FARPROC WRITENOTIFYPROC;
typedef FARPROC UNKNOWNITEMPROC;
typedef FARPROC ITEMREMOVEDPROC;
typedef FARPROC DISCONNECTPROC;
typedef FARPROC EVENTMSGPROC;
typedef FARPROC RATECHANGEPROC;
typedef FARPROC DEVICEREADPROC;

#endif

#ifdef __cplusplus
extern "C" {
#endif  /* __cplusplus */



/////////////////////////////////////////////////////////////////////////////////////////////
//
//	Initialization and Registration functions
//
/////////////////////////////////////////////////////////////////////////////////////////////
//	WTOPCsvrRevision()
//		Simply returns a version identifier for the DLL
//
//	SetThreadingModel()
//		Allows the application to select either the COINIT_MULTITHREADED
//		or COINIT_APARTMENTTHREADED model for the Server, 
//		(the default is COINIT_MULTITHREADED ).
//
//	InitWTOPCsvr()
//		Initializes DCOM, Security, and creates the Server ClassFactory
//		pCLSID_Svr points to a c-style GUID structure
//
//	ResetServerRate()
//		The specified ServerRate is the fastest rate at which the data can
//		be updated to a client.  The minimum server rate that may be selected is
//		10 ms.
//
//	SetVendorInfo()
//		Allows the application to specify a vendor specifc string to identify
//		the server.  This string will appear as part of the OPCSERVERSTATUS structure
//		returned by the GetStatus Interface.
//
//	UninitWTOPCsvr()
//		Uninitializes COM.
//
//	UpdateRegistry()
//		Makes the appropriate entries to the Windows Registry to identify the Server.
//		pCLSID_Svr points to a c-style GUID structure
//
//	AddLocalServiceKeysToRegistry()
//		Makes additional Registry Entries required if the Server is to be installed
//		as an NT Service.
//
//	UnregisterServer()
//		Removes the Registry Entries
//		pCLSID_Svr points to a c-style GUID structure
//
/////////////////////////////////////////////////////////////////////////////////////////////

__declspec(dllexport) WORD WINAPI WTOPCsvrRevision();

__declspec(dllexport) BOOL WINAPI SetThreadingModel(DWORD dwCoInit);

__declspec(dllexport) BOOL WINAPI InitWTOPCsvr (BYTE *pCLSID_Svr, UINT ServerRate);

__declspec(dllexport) BOOL WINAPI UninitWTOPCsvr ();

__declspec(dllexport) BOOL WINAPI ResetServerRate (UINT ServerRate);

__declspec(dllexport) BOOL WINAPI UpdateRegistry (BYTE *pCLSID_Svr, LPCSTR Name, LPCSTR Descr, LPCSTR ExePath);

__declspec(dllexport) BOOL WINAPI AddLocalServiceKeysToRegistry (LPCSTR Name);

__declspec(dllexport) BOOL WINAPI UnregisterServer (BYTE *pCLSID_Svr, LPCSTR Name);

__declspec(dllexport) void WINAPI SetVendorInfo (LPCSTR VendorInfo);


/////////////////////////////////////////////////////////////////////////////////////////////
//
//	OPC Item Functions
//
/////////////////////////////////////////////////////////////////////////////////////////////
//
// SetWtOPCsvrQualifier(...)
//		Allows the application to supply the delimiting character used to
//		seperate tag names in a hiearchial namespace.
//		(The delimiting character by default is '.')
//
//	CreateTag()
//		Add an OPC Item to the WtOPCsvr.DLL local tag list.  The dll takes care of
//		all client references to the tag and provides a callback to the application
//		if the tag is updated by a client.  Once a tag gets created, it's name will
//		automatically be presented in the browse list to any interested OPC Client.
//
//	UpdateTag()
//	UpdateTagWithTimeStamp()
//	UpdateTagByName()
//		Allows the Server Application to change the value, quality and timestamp of
//		a tag.  All updates are automatically provided to subscribing clients as defined
//		by the particular connection.  
//
//	SetTagProperties()
//		Tag Properties are values associated with an OPC Item other than its Value,
//		Quality and TimeStamp.  Any property value may be assigned by the server
//		for a defined tag.
//
//	ReadTag()
//	ReadTagWithTimeStamp()
//		Allows the Application to read each tag value from the WtOPCsvr.dll local Tag List.
//
//	SuspendTagUpdates()
//		When a Tag is created by the Server Application, it is automatically enabled for
//		client subscription.  The tag name will automatically be included in the server
//		browse list and any client may connect and read it's current value.  In certain 
//		applications, it is desirable to only create the tag whenever a client requests it.
//		The UNKNOWNITEMPROC callback may be used for dynamic tag allocation, but until the
//		tag is created, the tag name will not show up in the browse list.  Creation of the
//		tag followed by a call to SuspendTagUpdates() will allow the tag name to be browsed
//		by a client, but not subscribed to.  In this scenario, the WtOPCsvr.dll will issue
//		the UNKNOWNITEMPROC callback to allow the Application to enable the tag and begin
//		updating it's value only when being actively subscribed by an OPC Client.
//
//	RemoveTag()
//		Deletes a tag from the WtOPCsvr Tag List.
//
/////////////////////////////////////////////////////////////////////////////////////////////

__declspec(dllexport) char  WINAPI SetWtOPCsvrQualifier (char qualifier);

__declspec(dllexport) HANDLE WINAPI CreateTag (LPCSTR Name, VARIANT Value, WORD InitialQuality, BOOL IsWritable);

__declspec(dllexport) BOOL WINAPI UpdateTag (HANDLE TagHandle, VARIANT Value, WORD Quality);

__declspec(dllexport) BOOL WINAPI UpdateTagWithTimeStamp (HANDLE TagHandle, VARIANT Value, WORD Quality, FILETIME timestamp);

__declspec(dllexport) BOOL WINAPI UpdateTagByName (LPCSTR Name, VARIANT Value, WORD Quality);

__declspec(dllexport) BOOL WINAPI SetTagProperties (HANDLE TagHandle, DWORD PropertyID, LPCSTR Description, VARIANT Value);

__declspec(dllexport) BOOL WINAPI ReadTag (HANDLE TagHandle, VARIANT *pValue);

__declspec(dllexport) BOOL WINAPI ReadTagWithTimeStamp (HANDLE TagHandle, VARIANT *pValue, WORD *pQuality, FILETIME *pTimestamp);

__declspec(dllexport) BOOL WINAPI SuspendTagUpdates (HANDLE TagHandle, BOOL OnOff);

__declspec(dllexport) BOOL WINAPI RemoveTag (HANDLE TagHandle);

//RWD	2-Aug-2000	ARtI - Associates for Real-time Information
//-------------------------------------------------------------------------------------
// SetServerState	allows user to show problems
// SetHashing		enables hashing, sets suggested hashtable size
// UpdateTagToList	replaces UpdateTag call, requires prior StartUpdateTags else fails
//					and EndUpdateTags after all tags updated...
// variations UpdateTagWithTimeStampTo List and UpdateTagByNameToList could also be done...

__declspec(dllexport) void WINAPI SetServerState( OPCSERVERSTATE ServerState );

__declspec(dllexport) unsigned long WINAPI SetHashing( unsigned long sizeHash );

__declspec(dllexport) BOOL WINAPI StartUpdateTags ();

__declspec(dllexport) BOOL WINAPI UpdateTagToList (HANDLE TagHandle, VARIANT Value, WORD Quality);

__declspec(dllexport) BOOL WINAPI EndUpdateTags ();
//endRWD

/////////////////////////////////////////////////////////////////////////////////////////////
//
//	Auxilary Functions
//
/////////////////////////////////////////////////////////////////////////////////////////////
//
//	NumbrClientConnections()
//		Allows the Server Application to determine the number of OPC Clients
//		currently connected.
//
//	RequestDisconnect()
//		Provides the ability for the Server Application to request that all
//		Clients gracefully disconnect.  There is no guarantee that any client
//		will honor the request.
//
//	RefreshAllClients()
//		Forces the WtOPCsvr.DLL to issue data updates to all connected Clients.
//		(Primarily used for debugging.  Normal client updates are issued automatically
//		by the WtOPCsvr.DLL based on an iternal clock tick.)
//
// ConvertVBDateToFileTime()
// ConvertFileTimeToVBDate()
//		To be used with Visual Basic to convert between the OPC timestamp
//		iand a Date variable
//
////////////////////////////////////////////////////////////////////////////////////////////

__declspec(dllexport) int WINAPI NumbrClientConnections ();

__declspec(dllexport) void WINAPI RequestDisconnect ();

__declspec(dllexport) BOOL WINAPI RefreshAllClients();

__declspec(dllexport) BOOL WINAPI ConvertVBDateToFileTime1 (double *pVBDate, FILETIME *pFileTime);

__declspec(dllexport) BOOL WINAPI ConvertFileTimeToVBDate1 (FILETIME *pFileTime, double *pVBDate);

/////////////////////////////////////////////////////////////////////////////////////////////
//
//	Callback Enabling Functions
//
/////////////////////////////////////////////////////////////////////////////////////////////
__declspec(dllexport) BOOL WINAPI EnableWriteNotification (WRITENOTIFYPROC lpCallback, BOOL ConvertToNativeType);

__declspec(dllexport) BOOL WINAPI EnableUnknownItemNotification (UNKNOWNITEMPROC lpCallback);

__declspec(dllexport) BOOL WINAPI EnableItemRemovalNotification (ITEMREMOVEDPROC lpCallback);

__declspec(dllexport) BOOL WINAPI EnableDisconnectNotification (DISCONNECTPROC lpCallback);

__declspec(dllexport) BOOL WINAPI EnableEventMsgs (EVENTMSGPROC lpCallback);

__declspec(dllexport) BOOL WINAPI EnableRateNotification (RATECHANGEPROC lpCallback);

__declspec(dllexport) BOOL WINAPI EnableDeviceRead (DEVICEREADPROC lpCallback);

/////////////////////////////////////////////////////////////////////////////////////////////
//
// Support for Alarms & Events
// (March 2000)
//
/////////////////////////////////////////////////////////////////////////////////////////////
__declspec(dllexport) BOOL WINAPI UserAEMessage (LPCSTR Msg, DWORD Severity);

__declspec(dllexport) BOOL WINAPI UserAEMessageEx (ONEVENTSTRUCT Msg);

__declspec(dllexport) BOOL WINAPI UserAEMessageEx2 (DWORD hEventSubscription, ONEVENTSTRUCT Msg, BOOL bRefresh, BOOL bLastRefresh);

//
// Literal Definitions for LevelID
//
#define	ID_LOLO_LIMIT	1
#define ID_LO_LIMIT		2
#define ID_HI_LIMIT		3
#define ID_HIHI_LIMIT	4

__declspec(dllexport) BOOL WINAPI SetItemLevelAlarm (HANDLE TagHandle, int LevelID, float Limit, DWORD Severity, BOOL Enabled);

__declspec(dllexport) BOOL WINAPI GetItemLevelAlarm (HANDLE TagHandle, int LevelID, float *pLimit, DWORD *pSeverity, BOOL *pEnabled);




/////////////				Undocumented Functions						  ////////////////
//																						//
//				Undocumented function to Disable Demo Timer								//
//																						//
//////////////////////////////////////////////////////////////////////////////////////////


__declspec(dllexport) BOOL WINAPI Deactivate30MinTimer (LPCSTR Authorization);


/////////////////////////////////////////////////////////////////////////////////////////////
//
// Expanded Support for A&E callbacks
// (March 2001)
//
// The application can use the CWtAExCallback object definition below
// to generate overloaded functions to change the default behavior of the 
// WtOPCsvr.dll with regard to Alarms & Events.
//
// Basic support for Alarms & Events is supplied by the dll in the form of
// High & Low level alarms surrounding each defined OPC Item.  As the data values
// change, these level limits are compared and event messages generated as appropriate.
// The default behavior of the WtOPCsvr.dll does not provide for acknowledgements
// from an OPC Client or any filtering to be applied to the event subscription.
//
// By basing a C++ object on the CWtAExCallback object and overloading any of
// the virtual functions the application can expand upon or replace the basic
// A&E functionality of the dll as required.
/////////////////////////////////////////////////////////////////////////////////////////////

#ifdef _WTOPCSVRDLL
class __declspec(dllexport) CWtAExCallback
#else
class __declspec(dllimport) CWtAExCallback
#endif
{
public:
	// This is the object definition for the A&E callback object
   CWtAExCallback();
   virtual ~CWtAExCallback();

   // CIOPCEventServer::GetStatus(...)
   // The dll supplies the Last Update Time parameter to match the requesting subscriber.
   // The default implementation of this function returns valid parameters to the client.
   // Only overload this function if you have specific infprmation to add to the status structure
   virtual HRESULT GetStatus(OPCEVENTSERVERSTATUS** ppEventServerStatus, FILETIME LastUpdateTime);

   // CIOPCEventServer::QueryAvailableFilters(...)
   // The default implementation of this function returns zero for the FilterMask
   // ans S_OK, (i.e. Filters not supported!)
   virtual HRESULT QueryAvailableFilters (DWORD *pdwFilterMask);

   // CIOPCEventServer::QueryEventCategories (...)
   // The default implementation of this function supports two categories
   //		OPC_SIMPLE_EVENT	--  OPC_MSG_CAT:System Message
   //		OPC_CONDITION_EVENT	--	OPC_LEVEL_CAT:Level
   //							--	OPC_DISCRETE_CAT:Discrete
   virtual HRESULT QueryEventCategories( 
								DWORD		dwEventType,
								DWORD		*pdwCount,
								DWORD		**ppdwEventCategories,
								LPWSTR		**ppszEventCategoryDescs);

   // CIOPCEventServer::QueryConditionNames (...)
   // The default implementation of this function returns the following condition names
   //	OPC_MSG_CAT			--	"Informational Text"
   //	OPC_LEVEL_CAT		--	"Level Alarm"
   //	OPC_DISCRETE_CAT	--	"Discrete Alarm"
   virtual HRESULT QueryConditionNames( 
								DWORD		dwEventCategory,
								DWORD		*pdwCount,
								LPWSTR		**ppszConditionNames);

   // CIOPCEventServer::QuerySubConditionNames(...)
   // The default implementation of this function returns 
   // four subCondition Names for the "Level Alarm" condition:
   //		"Lo"
   //		"Lo Lo"
   //		"Hi"
   //		"Hi Hi"
   virtual HRESULT QuerySubConditionNames( 
								LPWSTR		szConditionName,
								DWORD		*pdwCount,
								LPWSTR		**ppszSubConditionNames);

   // CIOPCEventServer::QuerySourceConditions(...)
   // The default implementation of this function returns E_NOTIMPL
   virtual HRESULT QuerySourceConditions( 
								LPWSTR		szSource,
								DWORD		*pdwCount,
								LPWSTR		**ppszConditionNames);

  // CIOPCEventServer::QueryEventAttributes(...)
  // The default implementation of this function returns E_NOTIMPL
  virtual HRESULT QueryEventAttributes( 
								DWORD		dwEventCategory,
								DWORD		*pdwCount,
								DWORD		**ppdwAttrIDs,
								LPWSTR		**ppszAttrDescs,
								VARTYPE		**ppvtAttrTypes);

  // CIOPCEventServer::TranslateToItemIDs(...)
  // The default implementation of this function returns E_NOTIMPL
   virtual HRESULT TranslateToItemIDs( 
								LPWSTR		szSource,
								DWORD		dwEventCategory,
								LPWSTR		szConditionName,
								LPWSTR		szSubconditionName,
								DWORD		dwCount,
								DWORD		*pdwAssocAttrIDs,
								LPWSTR		**ppszAttrItemIDs,
								LPWSTR		**ppszNodeNames,
								CLSID		**ppCLSIDs);

  // CIOPCEventServer::GetConditionState(...)
  // The default implementation of this function returns E_NOTIMPL
   virtual HRESULT GetConditionState( 
								LPWSTR		szSource,
								LPWSTR		szConditionName,
								DWORD		dwNumEventAttrs,
								DWORD		*pdwAttributeIDs,
								OPCCONDITIONSTATE	**ppConditionState);

  // CIOPCEventServer::EnableConditionByArea(...)
  // The default implementation of this function returns E_NOTIMPL
   virtual HRESULT EnableConditionByArea( 
								DWORD			dwNumAreas,
								LPWSTR			*pszAreas);

  // CIOPCEventServer::EnableConditionBySource(...)
  // The default implementation of this function returns E_NOTIMPL
   virtual HRESULT EnableConditionBySource( 
								DWORD		dwNumSources,
								LPWSTR		*pszSources);

  // CIOPCEventServer::DisableConditionByArea(...)
  // The default implementation of this function returns E_NOTIMPL
   virtual HRESULT DisableConditionByArea( 
								DWORD		dwNumAreas,
								LPWSTR		*pszAreas);

  // CIOPCEventServer::DisableConditionBySource(...)
  // The default implementation of this function returns E_NOTIMPL
   virtual HRESULT DisableConditionBySource( 
								DWORD		dwNumSources,
								LPWSTR		*pszSources);

  // CIOPCEventServer::AckCondition(...)
  // The default implementation of this function returns S_FALSE
   virtual HRESULT AckCondition( 
								DWORD		dwCount,
								LPWSTR		szAcknowledgerID,
								LPWSTR		szComment,
								LPWSTR		*pszSource,
								LPWSTR		*pszConditionName,
								FILETIME	*pftActiveTime,
								DWORD		*pdwCookie,
								HRESULT		**ppErrors);

  // CIOPCEventServer::CreateAreaBrowser(...)
  // The default implementation of this function returns E_NOTIMPL
   virtual HRESULT CreateAreaBrowser( 
								REFIID		riid,
								LPUNKNOWN	*ppUnk);


  // CreateEventSubscription (DWORD hEventSubscription)
  // RemoveEventSubscription (DWORD hEventSubscription)
  // WtOPCsvr.dll will execute these funtions in the application callback
  // object whenevr a client subcribes and unsubscribes to the event server.
  // This allows the application to implement filters on an individual
  // subscription basis by calling UserAEMessageEx2.
   virtual HRESULT CreateEventSubscription (DWORD hEventSubscription, DWORD dwBufferTime, DWORD dwMaxSize, OPCHANDLE hClientSubscription); 
   virtual HRESULT RemoveEventSubscription (DWORD hEventSubscription); 

  // CIOPCEventSubscriptionMgt::SetFilter (...)
  // The default implementation of this function returns E_NOTIMPL
   virtual HRESULT SetEventSubscriptionFilter( 
							DWORD		hEventSubscription,
							DWORD		dwEventType,
							DWORD		dwNumCategories,
							DWORD		*pdwEventCategories,
							DWORD		dwLowSeverity,
							DWORD		dwHighSeverity,
							DWORD		dwNumAreas,
							LPWSTR		*pszAreaList,
							DWORD		dwNumSources,
							LPWSTR		*pszSourceList);

  // CIOPCEventSubscriptionMgt::GetFilter (...)
  // The default implementation of this function returns E_NOTIMPL
   virtual HRESULT GetEventSubscriptionFilter( 
							DWORD			hEventSubscription,
							DWORD			*pdwEventType,
							DWORD			*pdwNumCategories,
							DWORD			**ppdwEventCategories,
							DWORD			*pdwLowSeverity,
							DWORD			*pdwHighSeverity,
							DWORD			*pdwNumAreas,
							LPWSTR			**ppszAreaList,
							DWORD			*pdwNumSources,
							LPWSTR			**ppszSourceList);

  // CIOPCEventSubscriptionMgt::SelectReturnedAttributes (...)
  // The default implementation of this function returns E_NOTIMPL
   virtual HRESULT SelectEventSubscriptionReturnedAttributes( 
										DWORD		hEventSubscription,
										DWORD		dwEventCategory,
										DWORD		dwCount,
										DWORD		*dwAttributeIDs);
        
  // CIOPCEventSubscriptionMgt::GetReturnedAttributes (...)
  // The default implementation of this function returns E_NOTIMPL
   virtual HRESULT GetEventSubscriptionReturnedAttributes( 
										DWORD		hEventSubscription,
										DWORD		dwEventCategory,
										DWORD		*pdwCount,
										DWORD		**ppdwAttributeIDs);

  // CIOPCEventSubscriptionMgt::Refresh (...)
  // The default implementation of this function returns E_NOTIMPL
   virtual HRESULT RefreshEventSubscription(
							DWORD		hEventSubscription,
							DWORD		dwConnection);
        
  // CIOPCEventSubscriptionMgt::CancelRefresh (...)
  // The default implementation of this function returns E_NOTIMPL
   virtual HRESULT CancelEventSubscriptionRefresh(
							DWORD		hEventSubscription,
							DWORD		dwConnection);
        
  // CIOPCEventSubscriptionMgt::GetState (...)
  // The default implementation of this function returns E_NOTIMPL
   virtual HRESULT GetEventSubscriptionState( 
							DWORD			hEventSubscription,
							BOOL			*pbActive,
							DWORD			*pdwBufferTime,
							DWORD			*pdwMaxSize,
							OPCHANDLE		*phClientSubscription);
        
  // CIOPCEventSubscriptionMgt::SetState (...)
  // The default implementation of this function returns E_NOTIMPL
   virtual HRESULT SetEventSubscriptionState( 
							DWORD			hEventSubscription,
							BOOL			*pbActive,
							DWORD			*pdwBufferTime,
							DWORD			*pdwMaxSize,
							OPCHANDLE		hClientSubscription,
							DWORD			*pdwRevisedBufferTime,
							DWORD			*pdwRevisedMaxSize);


};

//
// Exported function to override the default CWtAExCallback object with one
// specific to the application.  Overload the functions you wish to process
// within your app.  If the application chooses to implement its own set
// of Alarms & Event categories, it should disable the internal limit checking
// for item values.
//
__declspec(dllexport) BOOL WINAPI SetAEServerCallback (BOOL DisableInternalItemLimits, CWtAExCallback *pCallback); 


#ifdef __cplusplus
}
#endif
