// (C) Copyright 2008-2009 by Autodesk, Inc. 
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted, 
// provided that the above copyright notice appears in all copies and 
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting 
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS. 
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC. 
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
//
// Use, duplication, or disclosure by the U.S. Government is subject to 
// restrictions set forth in FAR 52.227-19 (Commercial Computer
// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
// (Rights in Technical Data and Computer Software), as applicable.

//- IDynamicProperty2.h

#pragma once

typedef int PROPCAT;

using namespace System;
using namespace System::Runtime;
using namespace System::Runtime::InteropServices;
using namespace Autodesk::AutoCAD::Runtime;
using namespace Autodesk::AutoCAD::DatabaseServices;

#include "dynprops.h"

namespace Autodesk
{
  namespace AutoCAD
  {
    namespace Windows
    {
      namespace OPM
      {
        #pragma region IDynamicProperty2

          [InteropServices::Guid("9CAF41C2-CA86-4ffb-B05A-AC43C424D076")]
              [InteropServices::InterfaceTypeAttribute(InteropServices::ComInterfaceType::InterfaceIsIUnknown)]
              [InteropServices::ComVisible(true)]
              public interface class IDynamicProperty2
          {
              void GetGUID(
                  [InteropServices::Out] System::Guid% propGUID
              );
              void GetDisplayName(
                  [InteropServices::In,
                  InteropServices::Out,
                  InteropServices::MarshalAs(
                      InteropServices::UnmanagedType::BStr
                  )
                  ] interior_ptr<System::String^> name);
              void IsPropertyEnabled(
                  [InteropServices::In,
                  InteropServices::MarshalAs(
                      InteropServices::UnmanagedType::IUnknown
                  )
                  ] Object^ pUnk,
                  [InteropServices::Out] System::Int32% bEnabled
              );
              void IsPropertyReadOnly(
                  [InteropServices::Out] System::Int32% bReadonly
              );
              void GetDescription(
                  [InteropServices::In,
                  InteropServices::Out,
                  InteropServices::MarshalAs(
                      InteropServices::UnmanagedType::BStr
                  )
                  ] interior_ptr<System::String^> description
              );
              void GetCurrentValueName(
                  [InteropServices::In,
                  InteropServices::Out,
                  InteropServices::MarshalAs(
                      InteropServices::UnmanagedType::BStr
                  )
                  ] interior_ptr<System::String^> name
              );
              void GetCurrentValueType(
                  [InteropServices::Out] ushort% pVarType
              );
              void GetCurrentValueData(
                  [InteropServices::In,
                  InteropServices::MarshalAs(
                      InteropServices::UnmanagedType::IUnknown
                  )
                  ] Object^ pUnk,
                  [InteropServices::In,
                  InteropServices::Out,
                  InteropServices::MarshalAs(
                      InteropServices::UnmanagedType::Struct
                  )
                  ] interior_ptr<Object^> varData
              );
              void SetCurrentValueData(
                  [InteropServices::In,
                  InteropServices::MarshalAs(
                      InteropServices::UnmanagedType::IUnknown
                  )
                  ] Object^ pUnk,
                  [InteropServices::In,
                  InteropServices::MarshalAs(
                      InteropServices::UnmanagedType::Struct
                  )
                  ] Object^ varData
              );
              void Connect(
                  [InteropServices::In,
                  InteropServices::MarshalAs(
                      /*IDynamicPropertyNotify2*/
                      InteropServices::UnmanagedType::IUnknown
                  )
                  ] Object^ pSink
              );
              void Disconnect();
          };

        #pragma endregion

          
        #pragma region IDyanmicPropertyNotify2

          [InteropServices::Guid(
              "975112B5-5403-4197-AFB8-90C6CA73B9E1"
          )]
              [InteropServices::InterfaceTypeAttribute(
                  InteropServices::ComInterfaceType::InterfaceIsIUnknown
              )]
              [InteropServices::ComVisible(true)]
              public interface class IDynamicPropertyNotify2
          {
              void OnChanged(
                  [InteropServices::In,
                  InteropServices::MarshalAs(
                      InteropServices::UnmanagedType::IUnknown
                  )
                  ] Object^ pDynamicProperty
              );
              void GetCurrentSelectionSet(
                  [InteropServices::In,
                  InteropServices::Out,
                  InteropServices::MarshalAs(
                      InteropServices::UnmanagedType::Struct
                  )
                  ] interior_ptr<Object^> pSelection
              );
          };

        #pragma endregion



        #pragma region ICategorizeProperties

          [InteropServices::Guid("4D07FC10-F931-11ce-B001-00AA006884E5")]
              [InterfaceType(ComInterfaceType::InterfaceIsIUnknown)]
              [ComVisible(true)]
              public interface class ICategorizeProperties
          {
              void MapPropertyToCategory(
                  [In] int dispid,
                  [In, Out] interior_ptr<PROPCAT> ppropcat
              );

              void GetCategoryName(
                  [In] PROPCAT propcat,
                  [In] LCID lcid,
                  [In, Out, MarshalAs(UnmanagedType::BStr)] interior_ptr<String^> pbstrName
              );
          };

        #pragma endregion

        
        #pragma region IDynamicEnumProperty

          [InteropServices::Guid("8B384028-ACB1-11d1-A2B4-080009DC639A")]
              [InterfaceType(ComInterfaceType::InterfaceIsIUnknown)]
              [ComVisible(true)]
              public interface class IDynamicEnumProperty
          {
              void GetNumPropertyValues(
                  [In, Out] interior_ptr<int> numValues
              );

              void GetPropValueName(
                  [In] int index,
                  [In, Out, MarshalAs(UnmanagedType::BStr)] interior_ptr<String^> valueName
              );

              void GetPropValueData(
                  [In] int index,
                  [In, Out] interior_ptr<Object^> valueData
              );
          };

        #pragma endregion


        #pragma region IDynamicProperty

          [InteropServices::Guid("8B384028-ACA9-11d1-A2B4-080009DC639A")]
              [InteropServices::InterfaceTypeAttribute(
                  InteropServices::ComInterfaceType::InterfaceIsIUnknown)]
              [InteropServices::ComVisible(true)]

              public interface class IDynamicProperty
          {
              void GetGUID(
                  [InteropServices::Out] Guid% propGUID
              );

              void GetDisplayName(
                  [InteropServices::In,
                  InteropServices::Out,
                  InteropServices::MarshalAs(
                      InteropServices::UnmanagedType::BStr
                  )
                  ] String^% bstrName
              );

              void IsPropertyEnabled(
                  [InteropServices::In] IntPtr objectID,
                  [InteropServices::Out] bool% pbEnabled
              );

              void IsPropertyReadOnly(
                  [InteropServices::Out] bool% pbReadonly
              );

              void GetDescription(
                  [InteropServices::In,
                  InteropServices::Out,
                  InteropServices::MarshalAs(
                      InteropServices::UnmanagedType::BStr
                  )
                  ] String^% bstrName
              );

              void GetCurrentValueName(
                  [InteropServices::In,
                  InteropServices::Out,
                  InteropServices::MarshalAs(
                      InteropServices::UnmanagedType::BStr
                  )
                  ] String^% pbstrName
              );

              void GetCurrentValueType(
                  [InteropServices::Out] VARTYPE% pVarType
              );

              void GetCurrentValueData(
                  [InteropServices::In] IntPtr objectID,
                  [InteropServices::Out] Object^% pvarData
              );

              void SetCurrentValueData(
                  [InteropServices::In] IntPtr objectID,
                  [InteropServices::In] Object^ varData
              );

              void Connect(
                  [InteropServices::In,
                  InteropServices::MarshalAs(
                      InteropServices::UnmanagedType::IUnknown
                  )
                  ] Object^ pSink
              );

              void Disconnect();
          };

        #pragma endregion

        #pragma region IDynamicNotifyProperty

          [InteropServices::Guid("8B384028-ACA8-11d1-A2B4-080009DC639A")]
              [InteropServices::InterfaceTypeAttribute(
                  InteropServices::ComInterfaceType::InterfaceIsIUnknown
              )]
              [InteropServices::ComVisible(true)]

              public interface class IDynamicPropertyNotify
          {
              void OnChanged(
                  [InteropServices::In,
                  InteropServices::MarshalAs(
                      InteropServices::UnmanagedType::Interface
                  )
                  ] IDynamicProperty^ pDynamicProperty
              );

              void GetCurrentSelectionSet(
                  [InteropServices::Out,
                  InteropServices::MarshalAs(
                      InteropServices::UnmanagedType::Struct
                  )
                  ] Object^% pSelection
              );
          };

        #pragma endregion


      }
    }
  }
}