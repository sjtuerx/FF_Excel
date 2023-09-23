// Some copyright should be here...

using System;
using System.IO;
using UnrealBuildTool;

public class FF_Excel : ModuleRules
{
	public FF_Excel(ReadOnlyTargetRules Target) : base(Target)
	{
		PCHUsage = ModuleRules.PCHUsageMode.UseExplicitOrSharedPCHs;

        string ThirdPartyDir = "../Source/FF_Excel/ThirdParty";
        PrivateIncludePaths.Add(ThirdPartyDir);
        PrivateIncludePaths.Add(Path.Combine(ThirdPartyDir, "xlnt"));
        PrivateIncludePaths.Add(Path.Combine(ThirdPartyDir, "utfcpp"));

        PublicDependencyModuleNames.AddRange(
			new string[]
			{
				"Core",
				// ... add other public dependencies that you statically link with here ...
			}
			);
			
		PrivateDependencyModuleNames.AddRange(
			new string[]
			{
				"CoreUObject",
				"Engine",
				"Slate",
				"SlateCore",
                "DataRegistry",
				"Projects"
				// ... add private dependencies that you statically link with here ...	
			}
			);
		
		DynamicallyLoadedModuleNames.AddRange(
			new string[]
			{
				// ... add any modules that your module loads dynamically here ...
			}
			);
	}
}
