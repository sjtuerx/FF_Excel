// Copyright Epic Games, Inc. All Rights Reserved.

#pragma once

#include "Kismet/BlueprintFunctionLibrary.h"

THIRD_PARTY_INCLUDES_START
#include "xlnt.hpp"
THIRD_PARTY_INCLUDES_END

#include "FF_ExcelBPLibrary.generated.h"

UCLASS(BlueprintType)
class FF_EXCEL_API UFFExcel_Xlnt_Doc : public UObject
{
	GENERATED_BODY()

public:

	xlnt::workbook Document;

};

UCLASS(BlueprintType)
class FF_EXCEL_API UFFExcel_Xlnt_Worksheet : public UObject
{
	GENERATED_BODY()

public:

	xlnt::worksheet Worksheet;

};


USTRUCT(BlueprintType)
struct FF_EXCEL_API FXlntWorksheet
{
	GENERATED_BODY()

public:

	UPROPERTY(BlueprintReadWrite)
	UFFExcel_Xlnt_Worksheet* Worksheet_Object = nullptr;

	UPROPERTY(BlueprintReadWrite)
	FString Title;

	UPROPERTY(BlueprintReadWrite)
	int32 Id;

	UPROPERTY(BlueprintReadWrite)
	int32 Index;

	bool operator == (const FXlntWorksheet& Other) const
	{
		return Title == Other.Title && Id == Other.Id && Index == Other.Index;
	}

	bool operator != (const FXlntWorksheet& Other) const
	{
		return !(*this == Other);
	}
};

FORCEINLINE uint32 GetTypeHash(const FXlntWorksheet& Key)
{
	uint32 Hash_Title = GetTypeHash(Key.Title);
	uint32 Hash_Id = GetTypeHash(Key.Id);
	uint32 Hash_Index = GetTypeHash(Key.Index);

	uint32 GenericHash;
	FMemory::Memset(&GenericHash, 0, sizeof(uint32));
	GenericHash = HashCombine(GenericHash, Hash_Title);
	GenericHash = HashCombine(GenericHash, Hash_Id);
	GenericHash = HashCombine(GenericHash, Hash_Index);

	return GenericHash;
}

UDELEGATE(BlueprintAuthorityOnly)
DECLARE_DYNAMIC_DELEGATE_TwoParams(FDelegateXlntSheets, bool, IsSuccessfull, const TArray<FXlntWorksheet>&, Out_Sheets);

UCLASS()
class UFF_ExcelBPLibrary : public UBlueprintFunctionLibrary
{
	GENERATED_UCLASS_BODY()

	UFUNCTION(BlueprintCallable, meta = (DisplayName = "XLNT - Open Excel from File", Keywords = "xlnt, excel, open, file"), Category = "FF_Excel|xlnt|Document")
	static bool XLNT_Open_Excel_File(UFFExcel_Xlnt_Doc*& Out_Doc, FString In_Path, FString Password = "");

	UFUNCTION(BlueprintCallable, meta = (DisplayName = "XLNT - Open Excel from Memory", Keywords = "xlnt, excel, open, file"), Category = "FF_Excel|xlnt|Document")
	static bool XLNT_Open_Excel_Memory(UFFExcel_Xlnt_Doc*& Out_Doc, TArray<uint8> In_Buffer, FString FileName = "", FString Password = "");

	UFUNCTION(BlueprintCallable, meta = (DisplayName = "XLNT - Get All Worksheets", Keywords = "xlnt, excel, work, sheet, worksheet, get, all"), Category = "FF_Excel|xlnt|Worksheets")
	static void XLNT_Worksheets_Get_All(FDelegateXlntSheets DelegateSheets, UFFExcel_Xlnt_Doc* In_Doc);

	UFUNCTION(BlueprintPure, meta = (DisplayName = "XLNT - Get Worksheet Name", Keywords = "xlnt, excel, work, sheet, worksheet, get, name"), Category = "FF_Excel|xlnt|Worksheets")
	static bool XLNT_Worksheets_Get_Title(FString& Out_Name, UFFExcel_Xlnt_Worksheet* In_Sheet);

	UFUNCTION(BlueprintPure, meta = (DisplayName = "XLNT - Get Worksheet Id", Keywords = "xlnt, excel, work, sheet, worksheet, get, id"), Category = "FF_Excel|xlnt|Worksheets")
	static bool XLNT_Worksheets_Get_Id(int32& Out_Id, UFFExcel_Xlnt_Worksheet* In_Sheet);

	UFUNCTION(BlueprintPure, meta = (DisplayName = "XLNT - Get Worksheet Index", Keywords = "xlnt, excel, work, sheet, worksheet, get, index"), Category = "FF_Excel|xlnt|Worksheets")
	static bool XLNT_Worksheets_Get_Index(int32& Out_Index, UFFExcel_Xlnt_Worksheet* In_Sheet);

};
