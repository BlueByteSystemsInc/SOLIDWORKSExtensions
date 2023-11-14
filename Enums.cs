using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlueByte.SOLIDWORKS.Extensions.Enums
{
    public enum SheetMetalFeatureType_e
    {
        Unknown,
        SMBaseFlange,
        BreakCorner,
        CornerTrim,
        CrossBreak,
        EdgeFlange,
        FlatPattern,
        FlattenBends,
        Fold,
        FormToolInstance,
        Hem,
        Jog,
        LoftedBend,
        NormalCut,
        OneBend,
        ProcessBends,
        SheetMetal,
        SketchBend,
        SM3dBend,
        SMGusset,
        SMMiteredFlange,
        TemplateSheetMetal,
        ToroidalBend,
        UnFold

    }
    public enum swOpenError
        {
            [Description("No Error")]
            swNoError = 0,
            [Description("Unable to locate the file; the file is not loaded or the referenced file (that is, component) is suppressed")]
            swFileNotFoundError = 2,
            [Description(" A document with the same name is already open")]
            swFileWithSameTitleAlreadyOpen = 65536,
            [Description("The document was saved in a future version of SOLIDWORKS")]
            swFutureVersion = 8192,
            [Description("Another error was encountered (SOLIDWORKS Generic Error)")]
            swGenericError = 1,

            [Description("File type argument is not valid")]
            swInvalidFileTypeError = 1024,
            [Description("File encrypted by Liquid Machines")]
            swLiquidMachineDoc = 131072,
            [Description("File is open and blocked because the system memory is low, or the number of GDI handles has exceeded the allowed maximum")]
            swLowResourcesError = 262144,
            [Description("File contains no display data")]
            swNoDisplayData = 524288
        }
        public enum swOpenWarning
        {
            [Description("No Warning")]
            swNoError = 0,
            [Description("The document is already open. ")]
            swFileLoadWarning_AlreadyOpen = 128,
            [Description("The document was defined in the context of another existing document that is not loaded. ")]
            swFileLoadWarning_BasePartNotLoaded = 64,
            [Description("The document is opened silently and swOpenDocOptions_AutoMissingConfig is specified")]
            swFileLoadWarning_ComponentMissingReferencedConfig = 32768,
            [Description("Some dimensions are referenced to the models incorrectly; these dimensions are highlighted in red in the drawing.")]
            swFileLoadWarning_DimensionsReferencedIncorrectlyToModels = 16384,
            [Description("Warning appears if the internal ID of the document does not match the internal ID saved with the referencing document")]
            swFileLoadWarning_IdMismatch = 1,
            [Description("An attempt has been made to open an invisible document with a linked design table that was modified externally, and the design table cannot be updated because the document cannot be displayed; you must make the document visible to open it and update the design table.")]
            swFileLoadWarning_InvisibleDoc_LinkedDesignTableUpdateFail = 65536,
            [Description("The design table is missing")]
            swFileLoadWarning_MissingDesignTable = 131072,
            [Description("The document needs to be rebuilt")]
            swFileLoadWarning_NeedsRegen = 32,
            [Description("The document is read only")]
            swFileLoadWarning_ReadOnly = 2,
            [Description("The revolved feature dimensions were created in SOLIDWORKS 99 or earlier and are not synchronized with their corresponding dimensions in the sketch")]
            swFileLoadWarning_RevolveDimTolerance = 4096,
            [Description("The document is being used by another user")]
            swFileLoadWarning_SharingViolation = 4,
            [Description("The document is view only and a configuration other than the default configuration is set")]
            swFileLoadWarning_ViewOnlyRestrictions = 512



        }
    }

