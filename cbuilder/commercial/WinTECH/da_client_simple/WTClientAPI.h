
/////////////////////////////////////////////////////////////////////////////////////////////
//
//	Callback Definitions
//
//	Used by the DLL to pass control back to the Client Application
//	Each callback must be explicitly enabled by calling one of the following
//	exported API functions:
//				EnableOPCNotification (HANDLE hConnect, NOTIFYPROCAPI lpCallback);
//				EnableErrorNotification (ERRORPROCAPI lpCallback);
//				EnableShutdownNotification (HANDLE hConnect, SHUTDOWNPROCAPI lpCallback);
//				EnableClientEventMsgs (EVENTMSGPROC lpCallback);
//				EnableAECallback (AE_PROC lpCallback);
//				EnableExtendedAECallback (AE_PROC_EX lpCallback);
//
//		
//
//	NOTIFYPROCAPI
//		Signals the application when data has been updated by the server
//		prototype for the callback function is as follows:
//		void CALLBACK EXPORT OPCUpdateCallback (HANDLE hGroup, HANDLE hItem, VARIANT *pVar, FILETIME timestamp, DWORD quality)
//	ERRORPROCAPI
//		Signals the application when an error is detected by the dll.  If this callback
//		is not used  errors will generate a modal MessageBox which must be acknowledged
//		by the user)
//		prototype for the callback function is as follows:
//		void CALLBACK EXPORT ErrorMsgCallback (DWORD hResult, char *pMsg)
//		(the buffer supplied by the dll as pMsg is a temporary and data should be
//		copied to a permanant buffer by the application)
//	SHUTDOWNPROCAPI	
//		Signals the application if the connected server requests a disconnect;
//		The HANDLE parameter in the shutdown callback procedure identifies the connection.
//	EVENTMSGPROC
//		Establishes a callback to the application for displaying 
//		Debug DCOM Event Messages.  A character buffer is returned as a parameter
//		that contains a textual description of the message.
//	AE_PROC
//	AE_PROC_EX
//		Either callback may be enabled wih the difference being the amount of data
//		returned in the aragument list.  The prototypes for the two functions are as follows:
//		void CALLBACK EXPORT AECallback (HANDLE		hClientSubscription
//										char		*pSource,
//										FILETIME	EventTime,
//										char		*pDescription,
//										DWORD		dwSeverity);
//		void CALLBACK EXPORT ExtendedAECallback (HANDLE		hClientSubscription,
//												BOOL		bRefresh,
//												BOOL		bLastRefresh,
//												DWORD		dwCount,
//												ONEVENTSTRUCT	*pEvents)
//		notice that the Extended version of the callback contains the unfiltered parameter
//		list as returned by the A&E Server On_Event Interface.
//


#ifdef STRICT
typedef VOID (CALLBACK* NOTIFYPROCAPI)(HANDLE, HANDLE, VARIANT*, FILETIME, DWORD);
typedef VOID (CALLBACK* ERRORPROCAPI)(DWORD, CHAR*);
typedef VOID (CALLBACK* SHUTDOWNPROCAPI)(HANDLE);
typedef VOID (CALLBACK* EVENTMSGPROC)(CHAR*);
typedef VOID (CALLBACK* AE_PROC)(HANDLE, CHAR*, FILETIME, CHAR*, DWORD);
typedef VOID (CALLBACK* AE_PROC_EX)(HANDLE, DWORD, DWORD, DWORD, VOID*);
#else /* !STRICT */
typedef FARPROC NOTIFYPROCAPI;
typedef FARPROC ERRORPROCAPI;
typedef FARPROC SHUTDOWNPROCAPI;
typedef FARPROC EVENTMSGPROC;
typedef FARPROC AE_PROC;
typedef FARPROC AE_PROC_EX;
#endif


#ifdef __cplusplus
extern "C" {
#endif  /* __cplusplus */

//
// WTclientRevision()
//
// This function simply returns a WORD revision indication
//
__declspec(dllexport) WORD WINAPI WTclientRevision();

//
// WTclientCoInit()
//
// This function initializes DCOM using default security settings
//
__declspec(dllexport) BOOL WINAPI WTclientCoInit();

//
// NumberOfOPCServers(...)
//
// returns the number of available OPC servers 
// if UseOPCENUM is FALSE
//		the server list is obtained from the local Windows Registry
//		MachineName is ignored
//
// if UseOPCENUM is TRUE
//		the server list is obtained from the OPCENUM component as
//		supplied by OPC Foundation and MachineName may be used to 
//		obtain the list of servers from a remote machine
//
__declspec(dllexport) int  WINAPI NumberOfOPCServers (bool UseOPCENUM, LPCSTR MachineName);

//
// GetOPCServerName(...)
//
// Used to iterate through the server list obtained with NumberOfServers()
// User Buffer pointed to by pBuf is filled with the Server name at index of the Server List
// A returned value of FALSE indicates that the index is invalid.
//
__declspec(dllexport) BOOL  WINAPI GetOPCServerName (int index, char *pBuf, int BufSize);

//
// ConnectOPC (...)
//
// Establishes an OPC Connection withthe specified server
// INVALID_HANDLE_VALUE (-1) id returned if the connection cannot be established.
//
// if EnableDLLBuffering is TRUE
//		DLL will maintain a list of all items created by the application
//		and updated by the Server.  Application may issue ReadOPCItem() to
//		obtain new item data as desired.
//
// if EnableDLLBuffering is FALSE
//		DLL will not keep local copy of item data and subsequent calls to
//		ReadOPCItem() will fail.  User application must define proper callback
//		function, (EnableOPCNotification()), to obtain item updates
//
__declspec(dllexport) HANDLE  WINAPI ConnectOPC(LPCSTR MachineName, LPCSTR ServerName, BOOL EnableDLLBuffering); 


//
// ConnectOPC will always attempt an OPC 2.0 connection to the server if the
// Server supports the IConnectionPtContainer Interface.  The following function
// allows the client to force an OPC 1.0a connection using the IAdvise Interface.
// 
__declspec(dllexport) HANDLE  WINAPI ConnectOPC1(LPCSTR MachineName, LPCSTR ServerName, BOOL EnableDLLBuffering); 


//
// DisconnectOPC(...)
//
// Used to shutdown an OPC Connection
//
__declspec(dllexport) void  WINAPI DisconnectOPC(HANDLE hConnect);

//
// AddOPCGroup(...)
//
// Creates a new OPC Group for the defined connection
// Requested name, data rate and deadband specified in parameter list.
// Actual values for Rate and DeadBand are returned from the server in 
// the respective pointer locations.
// Groups are always created in the Active state.
// 
__declspec(dllexport) HANDLE  WINAPI AddOPCGroup (HANDLE hConnect, LPCSTR Name, DWORD *pRate, float *pDeadBand);

//
// RemoveOPCGroup(...)
//
// Removes and cleansup allocated resources for defined group
//
__declspec(dllexport) void  WINAPI RemoveOPCGroup (HANDLE hConnect, HANDLE hGroup);

//
// NumberOfOPCItems(...)
//
// Returns the number of OPC Items from the Browse Interface of the designated
// Server connection.  If the server does not support Browsing, a value of xero
// is returned.  This function fills an internal array of itemnames which may
// then be accessed via GetOPCItemName().
//
// This function is equivalent to calling BrowseItems using OPC_FLAT from
// the Root position.
//
__declspec(dllexport) int  WINAPI NumberOfOPCItems (HANDLE hConnect);

//
// GetOPCItemName(...)
//
// Allows user to iterate through the list of item names obtained from
// NumberOfOPCItems().
//

__declspec(dllexport) BOOL  WINAPI GetOPCItemName (HANDLE hConnect, int index, char *pBuf, int BufSize);

//
// GetOPCItemNameFromHandle(...)
//
// Allows user to iterate through the list of item names obtained from
// NumberOfOPCItems().
//

__declspec(dllexport) BOOL  WINAPI GetOPCItemNameFromHandle (HANDLE hConnect, HANDLE hGroup, HANDLE hItem, char *pBuf, int BufSize);

//
// GetOPCItemType(...)
//
// Allows the application to obtain the canonical data type and
// read/write access properties for a given item.
// A return value of FALSE, indicates that the requested item name does not
// exist in the connected Server.
//

__declspec(dllexport) BOOL  WINAPI GetOPCItemType (HANDLE hConnect, HANDLE hGroup, LPCSTR Name, VARTYPE *pType, DWORD *pAccessRights);

//
// AddOPCItem(...)
//
// Requests that the connected OPC Server add an item to the specified group.
// The return value identifies the item for future access by the user application.
// An INVALID_HANDLE_VALUE return, indicates that the requested item name does not
// exist in the connected Server.
//
__declspec(dllexport) HANDLE  WINAPI AddOPCItem (HANDLE hConnect, HANDLE hGroup, LPCSTR ItemName);

//
// AddOPCItemWithPath(...)
//
// Same as AddOPCItem except allows user to specify a Path.
//
__declspec(dllexport) HANDLE  WINAPI AddOPCItemWithPath (HANDLE hConnect, HANDLE hGroup, LPCSTR PathName, LPCSTR ItemName);

//
// RemoveOPCItem(...)
//
// Removes the specified OPC Item and cleans up resources
//
__declspec(dllexport) void  WINAPI RemoveOPCItem (HANDLE hConnect, HANDLE hGroup, HANDLE hItem);

//
// EnableOPCNotification(...)
//
// Establishes a callback function in the user application which will receive
// control when the value of an item is updated from the connected server.
//
// prototype for the callback function is as follows:
//		void CALLBACK EXPORT OPCUpdateCallback (HANDLE hGroup, HANDLE hItem, VARIANT *pVar, FILETIME timestamp, DWORD quality)
//
__declspec(dllexport) BOOL  WINAPI EnableOPCNotification (HANDLE hConnect, NOTIFYPROCAPI lpCallback);


//
// RefreshOPCGroup(...)
//
// Allows the Application to refresh the current value of all active items
// in an OPC Group definition
//
__declspec(dllexport) BOOL  WINAPI RefreshOPCGroup (HANDLE hConnect, HANDLE hGroup, DWORD Source);

//
// ChangeOPCGroupState(...)
//
// Allows the Application to activate and deactivate an OPC Group
//
__declspec(dllexport) BOOL  WINAPI ChangeOPCGroupState (HANDLE hConnect, HANDLE hGroup, BOOL Active);

//
// WriteOPCItem(...)
//
// Allows the controlling application to write to a defined OPC item
//
//
__declspec(dllexport) BOOL  WINAPI WriteOPCItem (HANDLE hConnect, HANDLE hGroup, HANDLE hItem, VARIANT *pVar, BOOL DoAsync);

//
// ReadOPCItem(...)
//
// May be used by an application to read the current value of an item.
// Is only valid if EnableDLLBuffering has been set to TRUE when the
// connection was established.
//
__declspec(dllexport) BOOL  WINAPI ReadOPCItem (HANDLE hConnect,  HANDLE hGroup, HANDLE hItem, VARIANT *pVar, FILETIME *pTimeStamp, DWORD *pQuality);

//
// ReadOPCItemFromDevice(...)
//
// Uses the SyncIO Interface to read an opc item directly from the Server.
//
__declspec(dllexport) BOOL  WINAPI ReadOPCItemFromDevice (HANDLE hConnect, HANDLE hGroup, HANDLE hItem, VARIANT *pVar, FILETIME *pTimeStamp, DWORD *pQuality);

//
// GetSvrStatus (...)
// 
// Allows the controlling application to interrogate the running
// status of an attached server.  pSvrStatus points to a structure containing
// a pointer to a buffer which is to receive a VendorInfo string (WSTR).
// VendorInfoBufSize defines the length of this buffer to keep the dll from overrunning.
// GetSvrStatus may be called with pSvrStatus = NULL, in which case a return
// value of TRUE indicates that the server processed the interface call, but 
// no data was returned.
//
__declspec(dllexport) BOOL WINAPI GetSvrStatus (HANDLE hConnect, OPCSERVERSTATUS *pSvrStatus, int VendorInfoBufSize);

//
// SetClientName (...)
// 
// Allows the controlling application to issue the OPC_Common SetClientName Interface
//
__declspec(dllexport) BOOL WINAPI SetClientName (HANDLE hConnect, LPCSTR Name);

//
// EnableErrorNotification(...)
//
// Establishes a callback function in the user application which will receive
// control when an error is detected by the dll.  If this callback is not used
// errors will generate a modal MessageBox which must be acknowledged by the user)
//
// prototype for the callback function is as follows:
//		void CALLBACK EXPORT ErrorMsgCallback (DWORD hResult, char *pMsg)
// (the buffer supplied by the dll as pMsg is a temporary and data should be
//  copied to a permanant buffer by the application)
//
__declspec(dllexport) BOOL  WINAPI EnableErrorNotification (ERRORPROCAPI lpCallback);


//
// EnableShutdownNotification(...)
//
// Establishes a callback to the application if the connected
// server requests a disconnect;
//
__declspec(dllexport) BOOL  WINAPI EnableShutdownNotification (HANDLE hConnect, SHUTDOWNPROCAPI lpCallback);


//
// EnableEventMsgs(...)
//
// Establishes a callback to the application for displaying 
// Debug DCOM Event Messages
//
__declspec(dllexport) BOOL WINAPI EnableClientEventMsgs (EVENTMSGPROC lpCallback);


/////////////					Extended Functions						  ////////////////
//								  (Added Aug2000)										//
//																						//
//																						//
//////////////////////////////////////////////////////////////////////////////////////////

//
// SetBrowseFilters(...)
//
// Allows the application to specify filters to use during Browse Operations.
// Item names may be filtered based on user defined string, data type, or
// Read/Write access rights
//
__declspec(dllexport) BOOL  WINAPI SetBrowseFilters (HANDLE hConnect, LPCSTR UserString, VARTYPE DataType, DWORD AccessType);

//
// GetNameSpace(...)
//
// returns OPC_NS_FLAT (2) or OPC_NS_HIERARCHIAL (1) for the specified server connection
//
__declspec(dllexport) BOOL  WINAPI GetNameSpace (HANDLE hConnect, WORD *pNameSpace);

//
// SetWTclientQualifier(...)
//
// Allows the application to supply the delimiting character used to
// seperate tag names in a hiearchial namespace.
// (The delimiting character by default is '.')
//
__declspec(dllexport) char  WINAPI SetWTclientQualifier (char qualifier);

//
// BrowseTo(...)
//
// Changes the current browse position for the server to the specified node.
// Use a NULL String or "Root" to browse to the top of the tree.
// NodeName should be a fully qualified node name as returned from 
// GetItemName
//
__declspec(dllexport) BOOL  WINAPI BrowseTo (HANDLE hConnect, LPCSTR NodeName);


//
// BrowseItemNames(...)
//
// Returns the number of items from the current browse position and fills
// the internal item name array with the node names that may be accessible 
// via GetItemName().
//
// The Filter parameter specifies:
//	OPC_BRANCH: returns only items that have children
//	OPC_LEAF: returns only items that don't have children
//	OPC_FLAT: Returns all OPC Item Names, (LEAFS ONLY), below the current position
//				(including all children of children).
// 
__declspec(dllexport) int  WINAPI BrowseItems (HANDLE hConnect, WORD Filter);

//
// SetItemHandle(...)
//
// Allows the application to set a reference identifier, (handle),.
// for an OPC Item to be used in the data update callback.  This
// allows the design to optimize the update cycle by making an item
// easier to find. The original Server Handle, (hItem), returned from
// AddOPCItem() is not changed.
//
__declspec(dllexport) BOOL  WINAPI SetItemUpdateHandle (HANDLE hConnect,  HANDLE hGroup, HANDLE hItem, HANDLE hUpdate);

//
// NumberOfItemProperties (...)
//
// Returns the number of item Properties for the specified item.
// An internal array of Property Descriptions is filled that may be accessed
// using the GetItemPropertyDescription function.  * * * CAUTION * * * there
// is only one item property array maintained by the dll for each server.
// The application must read all property descriptions for a given item 
// and save the PropertyID's, before calling NumberOfItemProperties for a new item.
//
//
__declspec(dllexport) int  WINAPI NumberOfItemProperties (HANDLE hConnect, LPCSTR ItemName);

//
// GetItemPropertyDescription(...)
//
// Returns the property values description for the last item specified 
// by NumberOfItemProperties().  The application may use the returned PropertyID
// to read the property value from the server using ReadPropertyValue().
//
__declspec(dllexport) BOOL  WINAPI GetItemPropertyDescription (HANDLE hConnect, int PropertyIndex, DWORD *pPropertyID, VARTYPE *pVT, BYTE *pDescr, int BufSize);
 
//
// ReadPropertyValue(...)
//
// Returns the property value for the item specified 
//
__declspec(dllexport) BOOL  WINAPI ReadPropertyValue (HANDLE hConnect, LPCSTR Itemname, DWORD PropertyID, VARIANT *pValue);





/////////////				Alarms & Events Functions					  ////////////////
//																						//
//																						//
//																						//
//////////////////////////////////////////////////////////////////////////////////////////

//
// NumberOfOPC_AEServers(...)
//
// returns the number of available OPC Alarms & Events servers 
//
//		The server list is obtained from the OPCENUM component as
//		supplied by OPC Foundation and MachineName may be used to 
//		obtain the list of servers from a remote machine
//
__declspec(dllexport) int  WINAPI NumberOfOPC_AEServers (LPCSTR MachineName);

//
// GetOPC_AEServerName(...)
//
// Used to iterate through the A&E server list obtained with NumberOf_AEServers()
// User Buffer pointed to by pBuf is filled with the Server name at index of the Server List
// A returned value of FALSE indicates that the index is invalid.
//
__declspec(dllexport) BOOL  WINAPI GetOPC_AEServerName (int index, char *pBuf, int BufSize);

//
// ConnectOPC_AE (...)
//
// Establishes an OPC Connection withthe specified A& Eserver
// INVALID_HANDLE_VALUE (-1) is returned if the connection cannot be established.
//
//
__declspec(dllexport) HANDLE  WINAPI ConnectOPC_AE(LPCSTR MachineName, LPCSTR ServerName); 


//
// DisconnectOPC_AE(...)
//
// Used to shutdown an OPC Connection
//
__declspec(dllexport) void  WINAPI DisconnectOPC_AE(HANDLE hConnect);

//
// Create_AE_Subscription
//
// Used to setup the A&E Server to begin sending event messages
//
__declspec(dllexport) BOOL  WINAPI Create_AE_Subscription (HANDLE hConnect, HANDLE SubscriptionHandle, DWORD *pBufferTime, DWORD *pMaxSize);

//
// Refresh_AE_Subscription
//
// Used to request a refresh from the attached A&E Server
//
__declspec(dllexport) BOOL  WINAPI Refresh_AE_Subscription (HANDLE hConnect, HANDLE SubscriptionHandle);

//
// EnableAECallback(...)
//
// Establishes a callback to the application for displaying 
// Alarms & Events Messages
//
__declspec(dllexport) BOOL WINAPI EnableAECallback (AE_PROC lpCallback);

//
// EnableExtendedAECallback(...)
//
// Establishes a callback to the application for displaying 
// Alarms & Events Messages (May be used in place of EnableAECallback
// to provide extended information returned from the server.)
//
__declspec(dllexport) BOOL WINAPI EnableExtendedAECallback (AE_PROC_EX lpCallback);

//
// GetAESvrStatus (...)
// 
// Allows the controlling application to interrogate the running
// status of an attached A&E server.  
//
__declspec(dllexport) BOOL WINAPI GetAESvrStatus (HANDLE hConnect, OPCEVENTSERVERSTATUS *pSvrStatus, int VendorInfoBufSize);

//
// The following functions are provided to allow the application
// to call directly into the Attached A&E Server.  The WTClient.DLL
// simply calls the associated Interface on the connected Server without
// modifying the parameter list in any way.
//
//	AckCondition(.. )
//	EnableConditionByArea(.. )
//	EnableConditionBySource(.. )
//	DisableConditionByArea(.. )
//	DisableConditionBySource(.. )
//	GetFilter (.. )
//	SetFilter (.. )
//
//	Refer to the OPC Alarms & Events Specification for detailed definition
//	of the parameter and return vallues.

__declspec(dllexport) HRESULT WINAPI AckCondition (	HANDLE hConnect, 
													DWORD		dwCount,
													LPWSTR		szAcknowledgerID,
													LPWSTR		szComment,
													LPWSTR		*pszSource,
													LPWSTR		*pszConditionName,
													FILETIME	*pftActiveTime,
													DWORD		*pdwCookie,
													HRESULT		**ppErrors);
__declspec(dllexport) HRESULT WINAPI EnableConditionByArea (HANDLE hConnect, DWORD dwNumAreas, LPWSTR	*pszAreas);
__declspec(dllexport) HRESULT WINAPI EnableConditionBySource (HANDLE hConnect, DWORD dwNumSources, LPWSTR	*pszSources);
__declspec(dllexport) HRESULT WINAPI DisableConditionByArea (HANDLE hConnect, DWORD dwNumAreas, LPWSTR *pszAreas);
__declspec(dllexport) HRESULT WINAPI DisableConditionBySource (HANDLE hConnect, DWORD	dwNumSources, LPWSTR *pszSources);
__declspec(dllexport) HRESULT WINAPI GetFilter(HANDLE hConnect, 
											HANDLE		hSubscription, 
											DWORD			*pdwEventType,
											DWORD			*pdwNumCategories,
											DWORD			**ppdwEventCategories,
											DWORD			*pdwLowSeverity,
											DWORD			*pdwHighSeverity,
											DWORD			*pdwNumAreas,
											LPWSTR			**ppszAreaList,
											DWORD			*pdwNumSources,
											LPWSTR			**ppszSourceList);
__declspec(dllexport) HRESULT WINAPI SetFilter(HANDLE hConnect, 
											HANDLE	hSubscription,
											DWORD		dwEventType,
											DWORD		dwNumCategories,
											DWORD		*pdwEventCategories,
											DWORD		dwLowSeverity,
											DWORD		dwHighSeverity,
											DWORD		dwNumAreas,
											LPWSTR		*pszAreaList,
											DWORD		dwNumSources,
											LPWSTR		*pszSourceList);

//
// ConvertFileTimeToVBDate (...)
//
// To be used with Visual Basic to convert the OPC timestamp
// into a Date variable
//
__declspec(dllexport) BOOL WINAPI ConvertFileTimeToVBDate (FILETIME *pFileTime, double *pVBDate);


/////////////				Undocumented Functions						  ////////////////
//																						//
//				Undocumented function to Disable Demo Timer								//
//																						//
//////////////////////////////////////////////////////////////////////////////////////////


__declspec(dllexport) BOOL  WINAPI Disable30MinTimer (LPCSTR Authorization);


#ifdef __cplusplus
}
#endif
