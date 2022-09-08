module DESign.BomTools.NmbsBom

open DESign.Helpers

type Tcx =
    {
        Length : FtIn
        TcxType : string
    }

type Bcx =
    {
        Length : FtIn
        BcxType : string
    }

type PunchedSeats =
    {
        Hl : FtIn option
        Hr : FtIn option
        Gage : float
    }

type Notes =
    {
        SpecialNotes : string option
        LoadNotes : string option
    }

type GirderGeometry =
    {
        A : FtIn
        N : int
        InteriorPanelLength : FtIn
        B : FtIn
    }

module Girder =

    type T =
        {
            Mark : string
            Quantity : int
            Size : string
            OverallLength : FtIn
            Tcxl : Tcx option
            Tcxr : Tcx option
            Bdl : float option
            Bdr : float option
            Bcxl : Bcx option
            Bcxr : Bcx option
            PunchedSeats : PunchedSeats option
            SeatSlopeLeft : float option
            SeatSlopeRight : float option
            Notes : Notes option
            GirderGeomtery : GirderGeometry
        }

