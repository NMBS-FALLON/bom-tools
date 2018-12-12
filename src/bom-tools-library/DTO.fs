module DESign.BomTools.DTO

type Load =
    {
        ID : string
        Type : string
        Category : string
        Position : string
        Load1Value : float
        Load1DistanceFt : float option
        Load1DistanceIn : float option
        Load2Value : float option
        Load2DistanceFt : float option
        Load2DistanceIn : float option
        Ref : string option
        LoadCases : string option
    }

type Joist =
    {
        Mark : string
        JoistSize : string
        LengthFt : float
        LengthIn : float
        LE_Depth : float
        RE_Depth : float
        OverallTC_Pitch : float
        Notes : string option
    }

type AdditionalJoistOnGirder =
    {
        LocationFt : float option
        LocationIn : float option
        LoadValue : float option
    }

type GirderPanel =
    {
        Number : int
        LengthFt : float
        LengthIn : float
    }

type GirderGeometry =
    {
        Mark : string
        A_Ft : float
        A_In : float
        B_Ft : float
        B_In : float
        Panels : GirderPanel list
    }
    
type Girder =
    {
        Mark : string
        GirderSize : string
        OverallLengthFt : float
        OverallLengthIn : float
        TCXL_LengthFt : float
        TCXR_LengthIn : float
        Notes : string option
        AdditionalJoists : AdditionalJoistOnGirder list
        GirderGemoetry : GirderGeometry
    }

type Note =
    {
        ID : string
        Note : string
    }

type Job  =
    {
        GeneralNotes : Note list
        ParticularNotes : Note list
        Loads : Load list
        Girders : Girder list
        Joists : Joist list
    }

