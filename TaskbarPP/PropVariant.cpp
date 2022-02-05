#define NOMINMAX
#include <PropIdlBase.h>
#include <propvarutil.h>
#include <cassert>
#include <span>
#include <vector>

class WinError
{
	HRESULT m_code;
public:
	 WinError(HRESULT code) : m_code{ code } {}

	static void ThrowIfError(HRESULT result)
	{
		if (result != S_OK)
			throw WinError{ result };
	}

	 operator HRESULT()
	{
		return m_code;
	}
};

class PropertyVariant
{
	PROPVARIANT m_value{};
	ULONG m_extractPosition{};	//for extracting values
public:
	PropertyVariant(bool value)
	{
		WinError::ThrowIfError(InitPropVariantFromBoolean(value, &m_value));
	}

	PropertyVariant(BOOL const* booleanVector, ULONG numElements)
	{
		WinError::ThrowIfError(InitPropVariantFromBooleanVector(booleanVector, numElements, &m_value));
	}

	template<typename T>
	PropertyVariant(T const* buffer, UINT bytes)
	{
		WinError::ThrowIfError(InitPropVariantFromBuffer(reinterpret_cast<const void*>(buffer), bytes));
	}

	template<typename T>
	PropertyVariant(T const& object) : PropertyVariant(&object, sizeof(object))
	{
	}

	PropertyVariant(REFCLSID clsid)
	{
		WinError::ThrowIfError(InitPropVariantFromCLSID(clsid, &m_value));
	}

	PropertyVariant(double value)
	{
		WinError::ThrowIfError(InitPropVariantFromDouble(value, &m_value));
	}

	PropertyVariant(double const* doubleVector, ULONG numElements)
	{
		WinError::ThrowIfError(InitPropVariantFromDoubleVector(doubleVector, numElements, &m_value));
	}

	PropertyVariant(FILETIME const& pFileTime)
	{
		WinError::ThrowIfError(InitPropVariantFromFileTime(&pFileTime, &m_value));
	}

	PropertyVariant(FILETIME const* fileTimeVector, ULONG numElements)
	{
		WinError::ThrowIfError(InitPropVariantFromFileTimeVector(fileTimeVector, numElements, &m_value));
	}

	enum class Type
	{
		Buffer,
		String,
		Vector,

	};

	PropertyVariant(REFGUID guid, Type as)
	{
		switch (as)
		{
		case PropertyVariant::Type::Buffer:
			WinError::ThrowIfError(InitPropVariantFromGUIDAsBuffer(guid, &m_value));
			break;
		case PropertyVariant::Type::String:
			WinError::ThrowIfError(InitPropVariantFromGUIDAsString(guid, &m_value));
			break;
		default:
			assert(false);
			break;
		}
	}

	PropertyVariant(SHORT value)
	{
		WinError::ThrowIfError(InitPropVariantFromInt16(value, &m_value));
	}

	PropertyVariant(SHORT const* int16Vector, ULONG numElements)
	{
		WinError::ThrowIfError(InitPropVariantFromInt16Vector(int16Vector, numElements, &m_value));
	}

	PropertyVariant(LONG value)
	{
		WinError::ThrowIfError(InitPropVariantFromInt32(value, &m_value));
	}

	PropertyVariant(LONG const* int32Vector, ULONG numElements)
	{
		WinError::ThrowIfError(InitPropVariantFromInt32Vector(int32Vector, numElements, &m_value));
	}

	PropertyVariant(LONGLONG value)
	{
		WinError::ThrowIfError(InitPropVariantFromInt64(value, &m_value));
	}

	PropertyVariant(LONGLONG const* int64Vector, ULONG numElements)
	{
		WinError::ThrowIfError(InitPropVariantFromInt64Vector(int64Vector, numElements, &m_value));
	}

	PropertyVariant(REFPROPVARIANT rhs, ULONG index)
	{
		WinError::ThrowIfError(InitPropVariantFromPropVariantVectorElem(rhs, index, &m_value));
	}

	PropertyVariant(HINSTANCE hInstance, UINT id)
	{
		WinError::ThrowIfError(InitPropVariantFromResource(hInstance, id, &m_value));
	}

	PropertyVariant(PCWSTR string)
	{
		WinError::ThrowIfError(InitPropVariantFromString(string, &m_value));
	}

	PropertyVariant(PCWSTR string, Type as)
	{
		if (as != Type::Vector)
			throw WinError{-1};

		WinError::ThrowIfError(InitPropVariantFromStringAsVector(string, &m_value));
	}

	PropertyVariant(PCWSTR* string, ULONG numElements)
	{
		WinError::ThrowIfError(InitPropVariantFromStringVector(string, numElements, &m_value));
	}

	PropertyVariant(STRRET* strRet, PCUITEMID_CHILD pidl)
	{
		WinError::ThrowIfError(InitPropVariantFromStrRet(strRet, pidl, &m_value));
	}

	PropertyVariant(USHORT value)
	{
		WinError::ThrowIfError(InitPropVariantFromUInt16(value, &m_value));
	}

	PropertyVariant(USHORT const* uint16Vector, ULONG numElements)
	{
		WinError::ThrowIfError(InitPropVariantFromUInt16Vector(uint16Vector, numElements, &m_value));
	}

	PropertyVariant(ULONG value)
	{
		WinError::ThrowIfError(InitPropVariantFromUInt32(value, &m_value));
	}

	PropertyVariant(ULONG const* uint32Vector, ULONG numElements)
	{
		WinError::ThrowIfError(InitPropVariantFromUInt32Vector(uint32Vector, numElements, &m_value));
	}

	PropertyVariant(ULONGLONG value)
	{
		WinError::ThrowIfError(InitPropVariantFromUInt64(value, &m_value));
	}

	PropertyVariant(ULONGLONG const* uint64Vector, ULONG numElements)
	{
		WinError::ThrowIfError(InitPropVariantFromUInt64Vector(uint64Vector, numElements, &m_value));
	}

	PropertyVariant(REFPROPVARIANT propvarSingle)
	{
		WinError::ThrowIfError(InitPropVariantVectorFromPropVariant(propvarSingle, &m_value));
	}

	template<Type type>
	bool Is() const
	{
		if (type == Type::String)
		{
			return IsPropVariantString(m_value);
		}
		else if (type == Type::Vector)
		{
			return IsPropVariantVector(m_value);
		}
	}

	auto compare(PropertyVariant const& rhs) const
	{
		return PropVariantCompare(m_value, rhs.m_value);
	}

	bool operator==(PropertyVariant const& rhs) const
	{
		return compare(rhs) == 0;
	}

	bool operator>(PropertyVariant const& rhs) const
	{
		return compare(rhs) == 1;
	}

	bool operator<(PropertyVariant const& rhs) const
	{
		return compare(rhs) == -1;
	}

	template<Type type>
	PropertyVariant as() const
	{

	}

	//Microsoft Docs issues
	PropertyVariant& operator>>(BOOL& value)
	{
		WinError::ThrowIfError(PropVariantGetBooleanElem(m_value, m_extractPosition++, &value));
		return *this;
	}

	PropertyVariant& operator>>(double& value)
	{
		WinError::ThrowIfError(PropVariantGetDoubleElem(m_value, m_extractPosition++, &value));
		return *this;
	}

	PropertyVariant& operator>>(PropertyVariant& rhs)
	{
		WinError::ThrowIfError(PropVariantGetElem(m_value, m_extractPosition++, &(rhs.m_value)));
		return *this;
	}

	PropertyVariant& operator>>(PROPVARIANT& rhs)
	{
		WinError::ThrowIfError(PropVariantGetElem(m_value, m_extractPosition++, &rhs));
		return *this;
	}

	auto size() const
	{
		return PropVariantGetElementCount(m_value);
	}

	PropertyVariant& operator>>(FILETIME& fileTime)
	{
		WinError::ThrowIfError(PropVariantGetFileTimeElem(m_value, m_extractPosition++, &fileTime));
		return *this;
	}

	PropertyVariant& operator>>(SHORT& value)
	{
		WinError::ThrowIfError(PropVariantGetInt16Elem(m_value, m_extractPosition++, &value));
		return *this;
	}

	PropertyVariant& operator>>(LONG& value)
	{
		WinError::ThrowIfError(PropVariantGetInt32Elem(m_value, m_extractPosition++, &value));
		return *this;
	}

	PropertyVariant& operator>>(LONGLONG& value)
	{
		WinError::ThrowIfError(PropVariantGetInt64Elem(m_value, m_extractPosition++, &value));
		return *this;
	}

	PropertyVariant& operator>>(PWSTR& string)
	{
		WinError::ThrowIfError(PropVariantGetStringElem(m_value, m_extractPosition++, &string));
		return *this;
	}

	PropertyVariant& operator>>(USHORT& value)
	{
		WinError::ThrowIfError(PropVariantGetUInt16Elem(m_value, m_extractPosition++, &value));
		return *this;
	}

	PropertyVariant& operator>>(ULONG& value)
	{
		WinError::ThrowIfError(PropVariantGetUInt32Elem(m_value, m_extractPosition++, &value));
		return *this;
	}

	PropertyVariant& operator>>(ULONGLONG& value)
	{
		WinError::ThrowIfError(PropVariantGetUInt64Elem(m_value, m_extractPosition++, &value));
		return *this;
	}

	PropertyVariant& setExtractPosition(ULONG newIndex)
	{
		m_extractPosition = newIndex;
		return *this;
	}

	operator BOOL() const
	{
		BOOL value{};
		WinError::ThrowIfError(PropVariantToBoolean(m_value, &value));
		return value;
	}

	PropertyVariant& operator>>(std::vector<BOOL>& span)
	{
		ULONG count{};
		WinError::ThrowIfError(PropVariantToBooleanVector(m_value, span.data(), span.size(), &count));
		if (count != span.size())
			span.shrink_to_fit();
		return *this;
	}

	operator BSTR() const
	{
		BSTR value;
		WinError::ThrowIfError(PropVariantToBSTR(m_value, &value));
		return value;
	}

	PropertyVariant& operator>>(std::span<unsigned char> buffer)
	{
		WinError::ThrowIfError(PropVariantToBuffer(m_value, buffer.data(), buffer.size_bytes()));
		return *this;
	}

	operator CLSID() const
	{
		CLSID value;
		WinError::ThrowIfError(PropVariantToCLSID(m_value, &value));
		return value;
	}

	operator double() const
	{
		double value;
		WinError::ThrowIfError(PropVariantToDouble(m_value, &value));
		return value;
	}

	operator GUID() const
	{
		GUID guid;
		WinError::ThrowIfError(PropVariantToGUID(m_value, &guid));
		return guid;
	}

	operator SHORT() const
	{
		SHORT value;
		WinError::ThrowIfError(PropVariantToInt16(m_value, &value));
		return value;
	}

	operator LONG() const
	{
		LONG int32Value;
		WinError::ThrowIfError(PropVariantToInt32(m_value, &int32Value));
		return int32Value;
	}

	operator LONGLONG() const
	{
		LONGLONG int64Value;
		WinError::ThrowIfError(PropVariantToInt64(m_value, &int64Value));
		return int64Value;
	}

	operator USHORT() const
	{
		USHORT uint16Value;
		WinError::ThrowIfError(PropVariantToUInt16(m_value, &uint16Value));
		return uint16Value;
	}

	operator ULONG() const
	{
		ULONG uint32Value;
		WinError::ThrowIfError(PropVariantToUInt32(m_value, &uint32Value));
		return uint32Value;
	}

	operator ULONGLONG() const
	{
		ULONGLONG uint64Value;
		WinError::ThrowIfError(PropVariantToUInt64(m_value, &uint64Value));
		return uint64Value;
	}
};

int main()
{
	BOOL value[] = { true, true, false, false };
	PropertyVariant p{ value, ARRAYSIZE(value) };
	
}