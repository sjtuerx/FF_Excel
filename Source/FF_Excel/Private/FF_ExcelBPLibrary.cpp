// Copyright Epic Games, Inc. All Rights Reserved.

#include "FF_ExcelBPLibrary.h"
#include "FF_Excel.h"

// UE Includes.
#include "Kismet/GameplayStatics.h"
#include "Misc/FileHelper.h"

UFF_ExcelBPLibrary::UFF_ExcelBPLibrary(const FObjectInitializer& ObjectInitializer)
: Super(ObjectInitializer)
{

}

bool UFF_ExcelBPLibrary::XLNT_Open_Excel_File(UFFExcel_Xlnt_Doc*& Out_Doc, FString In_Path, FString Password)
{
	if (In_Path.IsEmpty())
	{
		return false;
	}

	FString ExcelPath = FPlatformFileManager::Get().GetPlatformFile().ConvertToAbsolutePathForExternalAppForRead(*In_Path);
	if (!FPaths::FileExists(ExcelPath))
	{
		return false;
	}

	FString Extension = FPaths::GetExtension(ExcelPath, false);
	if (Extension != "xlsx")
	{
		return false;
	}

	TArray64<uint8> Temp_Buffer;
	FFileHelper::LoadFileToArray(Temp_Buffer, *ExcelPath, FILEREAD_AllowWrite);
	size_t Buffer_Size = Temp_Buffer.Num();

	std::vector<std::uint8_t> Load_Buffer(Buffer_Size);
	FMemory::Memcpy(Load_Buffer.data(), Temp_Buffer.GetData(), Buffer_Size);

	Out_Doc = NewObject<UFFExcel_Xlnt_Doc>();
	bool Result = false;

	if (Password.IsEmpty())
	{
		Result = Out_Doc->Document.load(Load_Buffer);
	}

	else
	{
		Result = Out_Doc->Document.load(Load_Buffer, TCHAR_TO_UTF8(*Password));
	}

	if (!Result)
	{
		UE_LOG(LogTemp, Warning, TEXT("Excel couldn't be opened from file : %s"), *FString(FPaths::GetCleanFilename(ExcelPath)));
		
		return false;
	}

	std::string Title;
	bool bHasTitle = Out_Doc->Document.has_title();
	if (bHasTitle)
	{
		Out_Doc->Document.title(Title);
	}

	UE_LOG(LogTemp, Display, TEXT("Excel opened from file - File Name : %s"), *FString(FPaths::GetCleanFilename(ExcelPath)));
	UE_LOG(LogTemp, Display, TEXT("Excel opened from file - Title : %s"), bHasTitle ? *FString(Title.c_str()) : *FString("TITLE_NONE"));

	return true;
}

bool UFF_ExcelBPLibrary::XLNT_Open_Excel_Memory(UFFExcel_Xlnt_Doc*& Out_Doc, TArray<uint8> In_Buffer, FString FileName, FString Password)
{
	if (In_Buffer.Num() == 0)
	{
		return false;
	}

	size_t Buffer_Size = In_Buffer.Num();
	std::vector<std::uint8_t> Load_Buffer(Buffer_Size);
	FMemory::Memcpy(Load_Buffer.data(), In_Buffer.GetData(), Buffer_Size);

	Out_Doc = NewObject<UFFExcel_Xlnt_Doc>();
	bool Result = false;

	if (Password.IsEmpty())
	{
		Result = Out_Doc->Document.load(Load_Buffer);
	}

	else
	{
		Result = Out_Doc->Document.load(Load_Buffer, TCHAR_TO_UTF8(*Password));
	}


	if (!Result)
	{
		UE_LOG(LogTemp, Warning, TEXT("Excel couldn't be opened from memory : %s"), *FString(FileName));

		return false;
	}

	std::string Title;
	bool bHasTitle = Out_Doc->Document.has_title();
	if (bHasTitle)
	{
		Out_Doc->Document.title(Title);
	}

	UE_LOG(LogTemp, Display, TEXT("Excel opened from memory - File Name : %s"), *FString(FileName));
	UE_LOG(LogTemp, Display, TEXT("Excel opened from meory - Title : %s"), bHasTitle ? *FString(Title.c_str()) : *FString("TITLE_NONE"));

	return true;
}

void UFF_ExcelBPLibrary::XLNT_Worksheets_Get_All(FDelegateXlntSheets DelegateSheets, UFFExcel_Xlnt_Doc* In_Doc)
{
	if (!IsValid(In_Doc))
	{
		return;
	}

	AsyncTask(ENamedThreads::AnyNormalThreadNormalTask, [DelegateSheets, In_Doc]()
		{
			TArray<FXlntWorksheetProperties> Array_Sheets;
			
			int32 SheetCount = In_Doc->Document.sheet_count();
			for (int32 Index_Sheet = 0; Index_Sheet < SheetCount; Index_Sheet++)
			{
				xlnt::worksheet Each_Worksheet = In_Doc->Document.sheet_by_index(Index_Sheet);

				if (!Each_Worksheet.is_valid())
				{
					continue;
				}
				
				FXlntWorksheetProperties Worksheet_Properties;
				Worksheet_Properties.Worksheet_Object = NewObject<UFFExcel_Xlnt_Worksheet>();
				Worksheet_Properties.Worksheet_Object->Worksheet = Each_Worksheet;
				Worksheet_Properties.Title = UTF8_TO_TCHAR(Each_Worksheet.title().c_str());
				Worksheet_Properties.Id = Each_Worksheet.id();
				Worksheet_Properties.Index = In_Doc->Document.index(Each_Worksheet);
				Array_Sheets.Add(Worksheet_Properties);
			}

			AsyncTask(ENamedThreads::GameThread, [DelegateSheets, Array_Sheets]()
				{
					DelegateSheets.ExecuteIfBound(true, Array_Sheets);
				}
			);

			return;
		}
	);
}

bool UFF_ExcelBPLibrary::XLNT_Worksheet_Get_By_Title(UFFExcel_Xlnt_Worksheet*& Out_Sheet, UFFExcel_Xlnt_Doc* In_Doc, FString In_Title)
{
	if (!IsValid(In_Doc))
	{
		return false;
	}

	xlnt::worksheet Found_Sheet = In_Doc->Document.sheet_by_title(TCHAR_TO_UTF8(*In_Title));
	if (Found_Sheet.is_valid())
	{
		Out_Sheet = NewObject<UFFExcel_Xlnt_Worksheet>();
		Out_Sheet->Worksheet = Found_Sheet;

		return true;
	}

	else
	{
		return false;
	}
}

bool UFF_ExcelBPLibrary::XLNT_Worksheet_Get_By_Id(UFFExcel_Xlnt_Worksheet*& Out_Sheet, UFFExcel_Xlnt_Doc* In_Doc, int32 In_Id)
{
	if (!IsValid(In_Doc))
	{
		return false;
	}

	xlnt::worksheet Found_Sheet = In_Doc->Document.sheet_by_id(In_Id);
	if (Found_Sheet.is_valid())
	{
		Out_Sheet = NewObject<UFFExcel_Xlnt_Worksheet>();
		Out_Sheet->Worksheet = Found_Sheet;

		return true;
	}

	else
	{
		return false;
	}
}

bool UFF_ExcelBPLibrary::XLNT_Worksheet_Get_By_Index(UFFExcel_Xlnt_Worksheet*& Out_Sheet, UFFExcel_Xlnt_Doc* In_Doc, int32 In_Index)
{
	if (!IsValid(In_Doc))
	{
		return false;
	}

	xlnt::worksheet Found_Sheet = In_Doc->Document.sheet_by_index(In_Index);
	if (Found_Sheet.is_valid())
	{
		Out_Sheet = NewObject<UFFExcel_Xlnt_Worksheet>();
		Out_Sheet->Worksheet = Found_Sheet;

		return true;
	}

	else
	{
		return false;
	}
}

bool UFF_ExcelBPLibrary::XLNT_Worksheet_Get_Title(FString& Out_Name, UFFExcel_Xlnt_Worksheet* In_Sheet)
{
	if (!IsValid(In_Sheet))
	{
		return false;
	}

	if (!In_Sheet->Worksheet.is_valid())
	{
		return false;
	}

	Out_Name = UTF8_TO_TCHAR(In_Sheet->Worksheet.title().c_str());
	return true;
}

bool UFF_ExcelBPLibrary::XLNT_Worksheet_Get_Id(int32& Out_Id, UFFExcel_Xlnt_Worksheet* In_Sheet)
{
	if (!IsValid(In_Sheet))
	{
		return false;
	}

	if (!In_Sheet->Worksheet.is_valid())
	{
		return false;
	}

	Out_Id = In_Sheet->Worksheet.id();
	return true;
}

bool UFF_ExcelBPLibrary::XLNT_Worksheet_Get_Index(int32& Out_Index, UFFExcel_Xlnt_Worksheet* In_Sheet)
{
	if (!IsValid(In_Sheet))
	{
		return false;
	}

	if (!In_Sheet->Worksheet.is_valid())
	{
		return false;
	}

	Out_Index = In_Sheet->Worksheet.workbook().index(In_Sheet->Worksheet);
	return true;
}

void UFF_ExcelBPLibrary::XLNT_Cells_At_Column(FDelegateXlntCells DelegateCells, UFFExcel_Xlnt_Worksheet* In_Sheet, int32 Index_Column)
{
	if (!IsValid(In_Sheet))
	{
		return;
	}

	if (!In_Sheet->Worksheet.is_valid())
	{
		return;
	}

	AsyncTask(ENamedThreads::AnyNormalThreadNormalTask, [DelegateCells, In_Sheet, Index_Column]()
		{
			TArray<FXlntCellProperties> Array_Cells;
			xlnt::row_t Row_Highest = In_Sheet->Worksheet.highest_row();
			xlnt::row_t Row_Lowest = In_Sheet->Worksheet.lowest_row();

			for (int32 Index_Row = (int32)Row_Lowest; Index_Row <= (int32)Row_Highest; ++Index_Row)
			{
				xlnt::cell_reference Each_Cell_Ref(Index_Column, Index_Row);
				xlnt::cell Each_Cell = In_Sheet->Worksheet.cell(Each_Cell_Ref);

				FXlntCellProperties EachCellProperties;
				EachCellProperties.Cell_Object = NewObject<UFFExcel_Xlnt_Cell>();
				EachCellProperties.Cell_Object->Cell = Each_Cell;
				EachCellProperties.Position_String = UTF8_TO_TCHAR(Each_Cell.column().column_string().c_str()) + FString::FromInt((int32)Each_Cell.row());
				EachCellProperties.Position_Referance = FVector2D((double)Each_Cell.column_index(), (double)Each_Cell.row());
	
				xlnt::cell_type CellType = Each_Cell.data_type();

				switch (CellType)
				{
				case xlnt::cell_type::empty:
					EachCellProperties.ValueType = EXlntDataTypes::Empty;
					break;
				case xlnt::cell_type::boolean:
					EachCellProperties.ValueType = EXlntDataTypes::Boolean;
					break;
				case xlnt::cell_type::date:
					EachCellProperties.ValueType = EXlntDataTypes::Date;
					break;
				case xlnt::cell_type::error:
					EachCellProperties.ValueType = EXlntDataTypes::Error;
					break;
				case xlnt::cell_type::inline_string:
					EachCellProperties.ValueType = EXlntDataTypes::String;
					break;
				case xlnt::cell_type::number:
					EachCellProperties.ValueType = EXlntDataTypes::Number;
					break;
				case xlnt::cell_type::shared_string:
					EachCellProperties.ValueType = EXlntDataTypes::String;
					break;
				case xlnt::cell_type::formula_string:
					EachCellProperties.ValueType = EXlntDataTypes::String;
					break;
				default:
					EachCellProperties.ValueType = EXlntDataTypes::None;
					break;
				}

				Array_Cells.Add(EachCellProperties);
			}
			
			AsyncTask(ENamedThreads::GameThread, [DelegateCells, Array_Cells]()
				{
					DelegateCells.ExecuteIfBound(true, Array_Cells);
				}
			);

			return;
		}
	);

}

bool UFF_ExcelBPLibrary::XLNT_Cell_Get_Value_Type(EXlntDataTypes& Out_Types, UFFExcel_Xlnt_Cell* In_Cell)
{
	if (!IsValid(In_Cell))
	{
		return false;
	}

	xlnt::cell_type CellType = In_Cell->Cell.data_type();
	
	switch (CellType)
	{
	case xlnt::cell_type::empty:
		Out_Types = EXlntDataTypes::Empty;
		break;
	case xlnt::cell_type::boolean:
		Out_Types = EXlntDataTypes::Boolean;
		break;
	case xlnt::cell_type::date:
		Out_Types = EXlntDataTypes::Date;
		break;
	case xlnt::cell_type::error:
		Out_Types = EXlntDataTypes::Error;
		break;
	case xlnt::cell_type::inline_string:
		Out_Types = EXlntDataTypes::String;
		break;
	case xlnt::cell_type::number:
		Out_Types = EXlntDataTypes::Number;
		break;
	case xlnt::cell_type::shared_string:
		Out_Types = EXlntDataTypes::String;
		break;
	case xlnt::cell_type::formula_string:
		Out_Types = EXlntDataTypes::String;
		break;
	default:
		Out_Types = EXlntDataTypes::None;
		break;
	}

	return true;
}

bool UFF_ExcelBPLibrary::XLNT_Cell_Get_Value_As_String(FString& Out_Value, UFFExcel_Xlnt_Cell* In_Cell)
{
	if (!IsValid(In_Cell))
	{
		return false;
	}

	Out_Value = UTF8_TO_TCHAR(In_Cell->Cell.to_string().c_str());
	return true;
}

bool UFF_ExcelBPLibrary::XLNT_Cell_Get_Value_As_Integer(int64& Out_Value, UFFExcel_Xlnt_Cell* In_Cell)
{
	if (!IsValid(In_Cell))
	{
		return false;
	}

	Out_Value = FMath::TruncToInt64(In_Cell->Cell.value<double>());
	return true;
}

bool UFF_ExcelBPLibrary::XLNT_Cell_Get_Value_As_Double(double& Out_Value, UFFExcel_Xlnt_Cell* In_Cell)
{
	if (!IsValid(In_Cell))
	{
		return false;
	}

	Out_Value = In_Cell->Cell.value<double>();
	return true;
}
