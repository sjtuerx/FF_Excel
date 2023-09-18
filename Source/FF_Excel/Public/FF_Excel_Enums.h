#pragma once

#include "CoreMinimal.h"

UENUM(BlueprintType)
enum class EXlntDataTypes : uint8
{
	None		UMETA(DisplayName = "None"),
	String		UMETA(DisplayName = "String"),
	Boolean		UMETA(DisplayName = "Boolean"),
	Number		UMETA(DisplayName = "Number"),
	Date		UMETA(DisplayName = "Date"),
	Error		UMETA(DisplayName = "Error"),
	Empty		UMETA(DisplayName = "Empty"),
};
ENUM_CLASS_FLAGS(EXlntDataTypes)