// TaskbarPP.cpp : Defines the entry point for the application.
//

#include "TaskbarPP.h"
#include <gtest/gtest.h>

TEST(Add, AddRecent)
{

}

TEST(Add, AddFrequent)
{

}

TEST(Add, AddCustom)
{

}

TEST(Add, AddTask)
{

}

TEST(Property, DisplayName)
{
	auto item = JumpListItem::CreateWithArguments(L"56", L"newTask3");
	EXPECT_EQ(item.DisplayName(), L"newTask3");
}

TEST(Property, EmptyDisplayName)
{
	auto item = JumpListItem::CreateWithArguments(L"78", L"");
	EXPECT_TRUE(item.DisplayName().empty());
}

TEST(Property, Arguments)
{

}

static void InitializeEnvironment()
{
	auto const initResult = CoInitializeEx(nullptr, COINIT_MULTITHREADED | COINIT_DISABLE_OLE1DDE);
	JumpList list;
	list += JumpListItem::CreateWithArguments(L"12", L"newTask");
	list += JumpListItem::CreateSeparator();
	list += JumpListItem::CreateWithArguments(L"34", L"newTask2");
	list.SaveAsync();

	ObjectArray<IApplicationDocumentLists> items[]{
		list.Items<JumpListGroupKind::Frequent>(),
		list.Items<JumpListGroupKind::Recent>(),
		list.Items<JumpListGroupKind::None>()
	};
	for (auto& item : items)
		std::cout << item.size() << '\n';

	std::cin.get();
}

int main(int argc, char** argv)
{
	InitializeEnvironment();
	::testing::InitGoogleTest(&argc, argv);
	return RUN_ALL_TESTS();
}

