// TaskbarPP.h : Include file for standard system include files,
// or project specific include files.

#pragma once

#include <Windows.h>

#include <string_view>
#include <vector>
#include <future>
#include <ShObjIdl_core.h>
#include <cassert>
#include <atlbase.h> //for CComPtr
#include <propkey.h> //for PKEY_Title
#include <propvarutil.h> //for InitPropVariantFrom...

#pragma comment(lib, "Propsys.lib")

using namespace std;

enum class JumpListItemKind
{
	Arguments,
	Separator
};

enum class JumpListGroupKind
{
	Frequent,
	None,
	Recent
};

class JumpListItem
{
	CComPtr<IShellLinkW> link;
	int m_index{};
	JumpListGroupKind kind = JumpListGroupKind::None;
	std::wstring m_group;
	
	auto getPropertyStore()
	{
		IPropertyStore* propertyStore{};
		link->QueryInterface(IID_PPV_ARGS(&propertyStore));
		assert(propertyStore);
		return propertyStore;
	}

	auto getId() const
	{
		PIDLIST_ABSOLUTE list{};
		link->GetIDList(&list);
		return list;
	}
public:
	JumpListItem()
	{
		link.CoCreateInstance(CLSID_ShellLink, nullptr);
	}

	std::wstring Arguments() const
	{
		std::wstring arguments(INFOTIPSIZE, {});
		link->GetArguments(&arguments[0], INFOTIPSIZE);
		return arguments;
	}

	void Arguments(std::wstring_view arguments)
	{
		link->SetArguments(arguments.data());
	}

	void Description(std::wstring_view descrption)
	{
		assert(descrption.size() <= INFOTIPSIZE);
		link->SetDescription(descrption.data());
	}

	std::wstring Description() const
	{
		std::wstring description(INFOTIPSIZE, {});
		link->GetDescription(&description[0], description.size());
		return description;
	}

	std::wstring DisplayName()
	{
		auto propertyStore = getPropertyStore();
		PROPVARIANT propvariant;
		propertyStore->GetValue(PKEY_Title, &propvariant);
		propertyStore->Commit();
		propertyStore->Release();
		return PropVariantToStringWithDefault(propvariant, L"");
	}

	void DisplayName(std::wstring_view displayName)
	{
		auto propertyStore = getPropertyStore();
		PROPVARIANT propvariant;
		InitPropVariantFromString(displayName.data(), &propvariant);
		propertyStore->SetValue(PKEY_Title, propvariant);
		propertyStore->Commit();
		propertyStore->Release();
	}

	std::wstring GroupName() const
	{

	}
	void GroupName(std::wstring_view groupName)
	{

	}

	JumpListItemKind Kind() const;

	std::wstring LogoPath() const
	{
		std::wstring path(MAX_PATH, {});
		int index{};
		link->GetIconLocation(
			&path[0],
			MAX_PATH,
			&index
		);
		assert(index == m_index);
		return path;
	}

	void LogoPath(std::wstring_view path)
	{
		//Might need GetShortPathNameW
		link->SetIconLocation(path.data(), m_index);
	}

	bool RemovedByUser() const
	{

	}


	static JumpListItem CreateSeparator()
	{
		JumpListItem item;
		auto propertyStore = item.getPropertyStore();
		PROPVARIANT propvariant;
		InitPropVariantFromBoolean(TRUE, &propvariant);
		propertyStore->SetValue(PKEY_AppUserModel_IsDestListSeparator, propvariant);
		propertyStore->Commit();
		propertyStore->Release();
		return item;
	}
	
	static JumpListItem CreateWithArguments(std::wstring_view arguments, std::wstring_view displayName)
	{
		JumpListItem item;

		{
			wchar_t path[MAX_PATH]{};
			GetModuleFileNameW(NULL, path, ARRAYSIZE(path));
			item.link->SetPath(path);
			item.link->SetArguments(arguments.data());
		}
		
		{
			auto propertyStore = item.getPropertyStore();
			PROPVARIANT propvariant;
			InitPropVariantFromString(displayName.data(), &propvariant);
			propertyStore->SetValue(PKEY_Title, propvariant);
			propertyStore->Commit();
			propertyStore->Release();
		}

		return item;
	}

	friend class JumpList;
};

template<typename Interface>
class ObjectArray
{
	CComPtr<IObjectArray> m_data;
	static inline const auto refiid = __uuidof(Interface);
public:
	ObjectArray(IObjectArray* data) : m_data(data)
	{
	}

	auto const size() const
	{
		UINT count{};
		m_data->GetCount(&count);
		return count;
	}

	auto operator[](UINT index) const
	{
		CComptr<Interface> value;
		m_data->GetAt(index, refiid, &value);
		return value;
	}
};

class JumpList
{
	CComPtr<IApplicationDocumentLists> plist;
//	std::vector<JumpListItem> newList;
	std::unordered_map<std::wstring, std::vector<JumpListItem>> newList;
public:
	JumpList()
	{
		plist.CoCreateInstance(CLSID_ApplicationDocumentLists);
	}

	template<JumpListGroupKind kind>
	auto Items() const
	{

		if constexpr (kind == JumpListGroupKind::Frequent)
		{
			IObjectArray* itemsRaw{};
			plist->GetList(
				ADLT_FREQUENT,
				10,	//0 to retrieve the full list
				IID_IObjectArray,
				reinterpret_cast<void**>(&itemsRaw)
			);

			return ObjectArray<IApplicationDocumentLists>(itemsRaw);
		}
		else if constexpr (kind == JumpListGroupKind::Recent)
		{
			IObjectArray* itemsRaw{};
			plist->GetList(
				ADLT_RECENT,
				10,	//0 to retrieve the full list
				IID_IObjectArray,
				reinterpret_cast<void**>(&itemsRaw)
			);

			return ObjectArray<IApplicationDocumentLists>(itemsRaw);
		}
		else if constexpr (kind == JumpListGroupKind::None)
		{
			IObjectArray* itemsRaw{};
			plist->GetList(
				ADLT_RECENT,
				10,	//0 to retrieve the full list
				IID_IObjectArray,
				reinterpret_cast<void**>(&itemsRaw)
			);

			return ObjectArray<IApplicationDocumentLists>(itemsRaw);
		}
	}

	JumpList& operator+=(JumpListItem&& item)
	{
		newList[item.GroupName()].push_back(std::move(item));
		//newList.push_back();
		return *this;
	}

	JumpListGroupKind SystemGroupKind() const;
	void SystemGroupKind(JumpListGroupKind kind);

	static bool IsSupported();
	static std::future<JumpList> LoadCurrentAsync();

	void Clear()
	{
		CComPtr<ICustomDestinationList> plist;
		auto hr = plist.CoCreateInstance(CLSID_DestinationList, nullptr);

		UINT numItems{};
		IObjectArray* removedItems;
		plist->BeginList(&numItems, IID_PPV_ARGS(&removedItems));
		plist->CommitList();
	}
	
	void SaveAsync()
	{
		CComPtr<ICustomDestinationList> plist;
		auto hr = plist.CoCreateInstance(CLSID_DestinationList, nullptr);

		UINT numItems{};
		IObjectArray* removedItems;
		plist->BeginList(&numItems, IID_PPV_ARGS(&removedItems));

		/*Build the list*/
		{
			CComPtr<IObjectCollection> pObjectCollection;
			pObjectCollection.CoCreateInstance(CLSID_EnumerableObjectCollection);
			for (auto const& item : newList)
			{
				pObjectCollection->AddObject(item.link.p);
			}

			CComPtr<IObjectArray> pArray;
			pObjectCollection->QueryInterface(IID_PPV_ARGS(&pArray));
			plist->AddUserTasks(pArray);
		}

		plist->CommitList();
	}
};

