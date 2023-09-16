#include "serialisation_utils.hpp"
#include "CoreMinimal.h"

std::string xlnt::serialize_double_to_string(double num)
{
	const FString str = FString::SanitizeFloat(num, 15);
	return TCHAR_TO_UTF8(*str);
}
