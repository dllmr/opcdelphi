#ifndef CUSTOMSINKS_H__
#define CUSTOMSINKS_H__

// Include all ATL stuff:
#ifdef __BORLANDC__
#include <atlvcl.h>
#endif
#include <atlbase.h>
#include <atlcom.h>

/////////////////////////////////////////////////////////////////////////////
// TCustomSink template class

/*

Notes on using the TCustomSink template class:

1- Define a COM object based on ATL, which should derive from
   CComObjectRootEx (or CComObjectRoot). You can use the ATL 
   object wizard to do this. 

2- The object should implement the events custom interface. Use 
   the wizard as properly (or typing) to add the corresponding code.

3- The object normally should be hidden and noncreatable in the IDL.
   Therefore, add those attributes to the coclass declared in the IDL.
   In code, comment (or erase) the object entry in the COM map, which 
   should be in the main file of your COM server, to avoid creation 
   through the Class Factory.

4- Declare the creatable sink as follows:

     typedef TCustomSink<CMySink, &IID_IMyEvents> CMyCreatableSink;
   
   CMySink in this case is the class you previously created using the
   wizard (or your typing skills), wich implements the custom interface 
   IMyEvents, meaning the custom event interface. 

5- Declare a member on your host class:
   
     CMyCreatableSink m_sink;

6- Use the members Connect and Disconnect as needed. When connected, the
   server will call, on the sink, the methods of the custom events interface
   the sink implements. That interface is actually an outgoing interface in the
   server.

7- Keep as a pattern having one sink per event interface. The name "Event Sink"
   suggest so.

*/

// 'Base' is the event sink class that derives from CComObjectRoot,
// and whatever interfaces the user wants to support on the object.
// piid is the address of the interface id which is implemented
// by the sink. Note that the base class must implement the
// custom interface for receiving events.
template <class Base, const IID& riid>
class TCustomSink : public Base
{
private:
   CComPtr<IUnknown> m_ptrSender; // Events sender
   DWORD m_dwCookie;              // Connection cookie

   // !!!! ionmun1 2000-01-05
   // For knowing whether the sink was created in the heap
   // or in the stack. This is a trick I did, but maybe
   // there is a better solution:
   union
   {
      #define HEAP_SIGNATURE __int64(0xA55AA55AA55AA55A)
      BOOL m_bHeap;
      __int64 m_llsignature;
   };

public:
   // !!!! ionmun1 2000-01-05
   // These operators have to do with the trick I described before:
   void* operator new(size_t aSize)
   {
      TCustomSink<Base, riid>* p =
         reinterpret_cast<TCustomSink<Base, riid>*>(::operator new(aSize));
      p->m_lsignature = HEAP_SIGNATURE;
      return p;
   }
   void operator delete(void* pThis, size_t aSize)
   {
      ::operator delete(pThis);
   }

public:
   TCustomSink() :
      m_dwCookie(0)
   {
      m_bHeap = (m_llsignature == HEAP_SIGNATURE);

      // Because it is created directly (without any class factory),
      // we increment the reference to the sink object
      AddRef();
   }

   ~TCustomSink()
   {
      // Disconnect from the server
      Disconnect();
      // Set refcount to 1 to protect destruction
      m_dwRef = 1L;
      FinalRelease();
   }

   // Implementation of IUnknown in the sink, to make it
   // instantiable:
   // ATL comment: if InternalAddRef or InteralRelease is
   // undefined then your class doesn't derive from CCustomSinkRoot
   STDMETHOD_(ULONG, AddRef)() { return InternalAddRef(); }
   STDMETHOD_(ULONG, Release)()
   {
      ULONG lRefCnt = InternalRelease();

      // If it was created in the heap, we destroy the object
      // if the reference count is zero:
      if (m_bHeap && (lRefCnt == 0))
      {
         delete this;
      }

      return lRefCnt;
   }

   // ATL comment: if _InternalQueryInterface is undefined then
   // you forgot BEGIN_COM_MAP
   STDMETHOD(QueryInterface)(REFIID iid, void ** ppvObject)
   {
      return _InternalQueryInterface(iid, ppvObject);
   }

public:
   // Methods for connecting/disconnecting from the event sender
   HRESULT __fastcall Connect(IUnknown* pSender)
   {
      HRESULT hResult = S_FALSE;

      if (pSender != m_ptrSender)
      {
         m_ptrSender = pSender;
         if (m_ptrSender != NULL)
         {
            // The "static_cast" is safe, because all COM objects
            // derive from IUnknown.
            hResult = AtlAdvise(m_ptrSender, static_cast<IUnknown*>(this),
                         riid,
                         &m_dwCookie);
         }
      }

      return hResult;
   }

   HRESULT __fastcall Disconnect()
   {
      HRESULT hResult = S_FALSE;

      if ( (m_ptrSender != NULL) &&
           (0 != m_dwCookie) )
      {
         hResult = AtlUnadvise(m_ptrSender,
                               riid,
                               m_dwCookie);
         m_dwCookie = 0;
         // Free the server:
         m_ptrSender = NULL;
      }

      return hResult;
   }
};

#endif // CUSTOMSINKS_H__
