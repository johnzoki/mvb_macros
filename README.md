# Sharepoint_Macros
A way to open Word templates from the company's Sharepoint using macros.

Documentation is in the code itself.
If possible prevent using umlauts in file or pathnames.

# Variabels you have to set before implementation.
havePresetFolder As Boolean

presetFolder As String

presetPath As String

## Examples:
### If havePresetFolder is True:
gavePresetFolder = True

presetFolder = "\Company\Documents\Presets\"

presetPath = "Preset_empty.dotm"

### If havePresetFolder is False:
gavePresetFolder = False

presetFolder = ""

presetPath = "Company\Documents\Presets\Preset_empty.dotm"
