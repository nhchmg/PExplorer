/*
 * Copyright 2003 Martin Fuchs
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 */


//
// Explorer clone
//
// shellbrowserimpl.cpp
//
// Martin Fuchs, 28.09.2003
//
// Credits: Thanks to Leon Finker for his explorer cabinet window example
//


#include <precomp.h>


HRESULT IShellBrowserImpl::QueryInterface(REFIID iid, void **ppvObject)
{
    if (!ppvObject)
        return E_POINTER;

    if (iid == IID_IUnknown)
        *ppvObject = (IUnknown *)static_cast<IShellBrowser *>(this);
    else if (iid == IID_IOleWindow)
        *ppvObject = static_cast<IOleWindow *>(this);
    else if (iid == IID_IShellBrowser)
        *ppvObject = static_cast<IShellBrowser *>(this);
    else if (iid == IID_ICommDlgBrowser)
        *ppvObject = static_cast<ICommDlgBrowser *>(this);
    else if (iid == IID_IServiceProvider)
        *ppvObject = static_cast<IServiceProvider *>(this);
    else {
        *ppvObject = NULL;
        return E_NOINTERFACE;
    }

    return S_OK;
}

HRESULT IShellBrowserImpl::QueryService(REFGUID guidService, REFIID riid, void **ppvObject)
{
    if (!ppvObject)
        return E_POINTER;

    ///@todo use guidService

    if (riid == IID_IUnknown)
        *ppvObject = (IUnknown *)static_cast<IShellBrowser *>(this);
    else if (riid == IID_IOleWindow)
        *ppvObject = static_cast<IOleWindow *>(this);
    else if (riid == IID_IShellBrowser)
        *ppvObject = static_cast<IShellBrowser *>(this);
    else if (riid == IID_ICommDlgBrowser)
        *ppvObject = static_cast<ICommDlgBrowser *>(this);
    else if (riid == IID_IServiceProvider)
        *ppvObject = static_cast<IServiceProvider *>(this);
    else if (riid == IID_IOleCommandTarget)
        *ppvObject = static_cast<IOleCommandTarget *>(this);
    else {
        *ppvObject = NULL;
        return E_NOINTERFACE;
    }

    return S_OK;
}

HRESULT IShellBrowserImpl::QueryStatus(const GUID *pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT *pCmdText)
{
    return E_FAIL;  ///@todo implement IOleCommandTarget
}

HRESULT IShellBrowserImpl::Exec(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvaIn, VARIANT *pvaOut)
{
    return E_FAIL;  ///@todo implement IOleCommandTarget
}


// process default command: look for folders and traverse into them
/*
HRESULT IShellBrowserImpl::OnDefaultCommand(IShellView *ppshv)
{
    IDataObject *selection;

    HRESULT hr = ppshv->GetItemObject(SVGIO_SELECTION, IID_IDataObject, (void **)&selection);
    if (FAILED(hr))
        return hr;

    PIDList pidList;

    hr = pidList.GetData(selection);
    if (FAILED(hr)) {
        selection->Release();
        return hr;
    }

    hr = OnDefaultCommand(pidList);

    selection->Release();

    return hr;
}
*/
/*
HRESULT IShellBrowserImpl::OnDefaultCommand(IShellView *ppshv)
{
	
	if (!ppshv)
		return E_INVALIDARG;

	IDataObject* pDataObj = NULL;
	HRESULT hr = ppshv->GetItemObject(SVGIO_SELECTION, IID_IDataObject, (void**)&pDataObj);
	if (SUCCEEDED(hr))
	{
		// 获取当前选中项的绝对PIDL
		FORMATETC formatEtc = { (CLIPFORMAT)RegisterClipboardFormat(CFSTR_SHELLIDLIST), nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
		STGMEDIUM medium = {0};
		hr = pDataObj->GetData(&formatEtc, &medium);
		if (SUCCEEDED(hr))
		{
			LPIDA pIDA = reinterpret_cast<LPIDA>(GlobalLock(medium.hGlobal));
			if (pIDA)
			{
				// 获取选中项的绝对PIDL
				LPCITEMIDLIST pidlParent = reinterpret_cast<LPCITEMIDLIST>((LPBYTE)pIDA + pIDA->aoffset[0]);
				LPCITEMIDLIST pidlChild = reinterpret_cast<LPCITEMIDLIST>((LPBYTE)pIDA + pIDA->aoffset[1]);
				PIDLIST_ABSOLUTE pidlSelected = ILCombine(pidlParent, pidlChild);

				if (pidlSelected)
				{
					// 使用ShellExecuteEx打开并选中项
					SHELLEXECUTEINFO sei = { 0 };
					sei.cbSize = sizeof(SHELLEXECUTEINFO);
					sei.fMask = SEE_MASK_IDLIST;
					sei.lpIDList = pidlSelected;
					sei.nShow = SW_SHOWNORMAL;

					if (ShellExecuteEx(&sei))
					{
						hr = S_OK;
					}
					else
					{
						hr = HRESULT_FROM_WIN32(GetLastError());
					}
					CoTaskMemFree(pidlSelected);
				}
				GlobalUnlock(medium.hGlobal);
			}
			ReleaseStgMedium(&medium);
		}
		pDataObj->Release();
	}
	return hr;
}

*/
HRESULT IShellBrowserImpl::OnDefaultCommand(IShellView *ppshv)
{
	if (!ppshv)
		return E_INVALIDARG;

	IDataObject* pDataObj = nullptr;
	HRESULT hr = ppshv->GetItemObject(SVGIO_SELECTION, IID_IDataObject, (void**)&pDataObj);
	if (SUCCEEDED(hr))
	{
		FORMATETC formatEtc = { (CLIPFORMAT)RegisterClipboardFormat(CFSTR_SHELLIDLIST), nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
		STGMEDIUM medium = { 0 };
		hr = pDataObj->GetData(&formatEtc, &medium);
		if (SUCCEEDED(hr))
		{
			LPIDA pIDA = reinterpret_cast<LPIDA>(GlobalLock(medium.hGlobal));
			if (pIDA)
			{
				LPCITEMIDLIST pidlParent = reinterpret_cast<LPCITEMIDLIST>((LPBYTE)pIDA + pIDA->aoffset[0]);
				LPCITEMIDLIST pidlChild = reinterpret_cast<LPCITEMIDLIST>((LPBYTE)pIDA + pIDA->aoffset[1]);
				PIDLIST_ABSOLUTE pidlSelected = ILCombine(pidlParent, pidlChild);

				if (pidlSelected)
				{
					IShellFolder* pShellFolder = nullptr;
					hr = SHBindToParent(pidlSelected, IID_IShellFolder, (void**)&pShellFolder, nullptr);
					if (SUCCEEDED(hr))
					{
						PCUITEMID_CHILD pidlRelSelected = ILFindLastID(pidlSelected);

						IContextMenu* pContextMenu = nullptr;
						hr = pShellFolder->GetUIObjectOf(NULL, 1, &pidlRelSelected, IID_IContextMenu, nullptr, (void**)&pContextMenu);
						if (SUCCEEDED(hr))
						{
							HMENU hMenu = CreatePopupMenu();
							if (hMenu)
							{
								hr = pContextMenu->QueryContextMenu(hMenu, 0, FCIDM_SHVIEWFIRST, FCIDM_SHVIEWLAST, CMF_DEFAULTONLY);
								if (SUCCEEDED(hr))
								{
									UINT uDefaultCmd = 0;
									BOOL bFoundDefault = FALSE;

									UINT uCount = GetMenuItemCount(hMenu);
									for (UINT i = 0; i < uCount; ++i)
									{
										//TCHAR szMenuText[256] = { 0 }; // 分配一个足够大的缓冲区存储菜单项文本
										MENUITEMINFO mii = { 0 };
										mii.cbSize = sizeof(MENUITEMINFO);
										mii.fMask = MIIM_ID | MIIM_STATE;// | MIIM_TYPE;
										
										//mii.dwTypeData = szMenuText; // 指定要接收菜单项文本的缓冲区
										//mii.cch = sizeof(szMenuText) / sizeof(TCHAR); // 指定缓冲区的字符数
										if (GetMenuItemInfo(hMenu, i, TRUE, &mii))
										{
											if (mii.fState & MFS_DEFAULT)
											{
												uDefaultCmd = mii.wID;
												bFoundDefault = TRUE;
												break;
											}
										}
									}

									if (bFoundDefault)
									{
										CMINVOKECOMMANDINFO cmi = { 0 };
										cmi.cbSize = sizeof(CMINVOKECOMMANDINFO);
										cmi.lpVerb = MAKEINTRESOURCEA(uDefaultCmd);
										cmi.nShow = SW_SHOWNORMAL;
										hr = pContextMenu->InvokeCommand(&cmi);
									}
								}
								DestroyMenu(hMenu);
							}
							else
							{
								hr = E_FAIL;
							}
							pContextMenu->Release();
						}
						pShellFolder->Release();
					}
					CoTaskMemFree(pidlSelected);
				}
				GlobalUnlock(medium.hGlobal);
			}
			ReleaseStgMedium(&medium);
		}
		pDataObj->Release();
	}
	return hr;
}