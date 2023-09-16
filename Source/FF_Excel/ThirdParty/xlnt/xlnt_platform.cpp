
#include "CoreMinimal.h"
#include "xlnt_platform.hpp"

namespace xlnt_custom
{
	std::string to_string(uint32 value)
	{
		FString Number = FString::FromInt(value);
		return TCHAR_TO_UTF8(*Number);
	}
}