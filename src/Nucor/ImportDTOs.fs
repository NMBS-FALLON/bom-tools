module DESign.BomTools.Vulcraft.Import.Dtos

open DESign.BomTools
open DESign.Helpers

type ImportDto = 


  {
      Mark : string option
      Quantity : string option
      DesignationInfo : {| Depth : string option ; DesignationType : string option ; Loading : string option |}
      OverallLength : {| Feet : string option ; Inch : string option |}
      Tcxl : {| Length : {| Feet : string option ; Inch : string option |} ; TcxType : string option|}
      Tcxr : {| Length : {| Feet : string option ; Inch : string option |} ; TcxType : string option|}
      Bcxl : {| Length : {| Feet : string option ; Inch : string option |} ; BcxType : string option|}
      Bcxr : {| Length : {| Feet : string option ; Inch : string option |} ; BcxType : string option|}
      BearingDepthLe : string option
      BearingDepthRe : string option
      BaySlope : string option
      BaySlopeHighEnd : string option
      LeSeatSlope : string option
      LeSeatSlopeHighLow : string option
      ReSeatSlope : string option
      ReSeatSlopeHighLow : string option
      LeBearingSlotInfo :
        {|
            Location : {|Feet : string option ; Inch : string option |}
            Size : string option
            Length : string option
            Gage : string option
        |}
      ReBearingSlotInfo :
        {|
            Location : {|Feet : string option ; Inch : string option |}
            Size : string option
            Length : string option
            Gage : string option
        |}
      NetUplift : string option
      AdLoad : string option
      Remarks : string option
      GeneralRemarks : string option
      TcConcentratedLoads : {| Load : string option ; DistanceFt : string option ; DistanceIn : string option|} list
      BcConcentratedLoads : {| Load : string option ; DistanceFt : string option ; DistanceIn : string option|} list
      BcUniformLoad : string option
      TcBendCheckLoad : string option
      BcBendCheckLoad : string option
      MainDrifts :
        {|
            StartLocFt : string option
            StartLocIn : string option
            StartLoad : string option
            EndLoadCol : string option
            LengthFt : string option
            LengthIn : string option
        |} list
  }


type Tcx = 
    {
        Length : FtIn
        TcxType : string
    }

type Bcx =
    {
        Length : FtIn
        BcxType : int
    }

type SlopeInfo =
    {
        BaySlope : float option
        HighEnd : string option
        LeSeatSlope : float option
        LeHighLow : string option
        ReSeatSlope : float option
        ReHighLow : string option
    }

type SeatSlot =
    {
        Location : FtIn
        Side : string option
        Size : float
        Length : float
        Gage : float
    }

type TopChordHoles =
    {
        Size : int
        Location : FtIn
        Side : string option
        DistanceBetweenHoles : float
        Gage : float
    }

type GirderPanels =
    {
        A : FtIn
        InterirorNumSpaces : int
        InteriorSpace : FtIn
        B : FtIn
    }

type ConcentratedLoad =
    {
        Load : string
        Distance : FtIn
    }

type UniformLoad = 
    {
        StartLocation : FtIn
        StartLoad : float 
        EndLocation : FtIn  
        EndLoad : float
    }

type LateralMoment =
    {
        MomentType : string
        Left : float option
        Right : float option
    }

type AxialLoad =
    {
        TcOrBc : string option
        Wind : float option
        Seismic : float option
        SeismicEm : float option
    }

type MethodOfAxialTransfer =
    {
        TcOrBc : string
        Method : string
        Detail : int
        KinfePlate : string
    }


type Girder = 
    {
        Mark : string
        Quantity: int
        Depth: int
        DesignationType: string
        NumPanels : int
        Load: string
        OverallLengthFt : FtIn
        Tcxl : Tcx option
        Tcxr : Tcx option
        Bcxl : Bcx option
        Bcxr : Bcx option
        BrgDepthLeft : float
        BrgDepthRight : float
        SlopeInfo : SlopeInfo
        LeSeatSlots : SeatSlot option
        ReSeatSlots : SeatSlot option
        LeTcHoles : TopChordHoles option
        ReTcHoles : TopChordHoles option
        NetUplift : string
        AdLoad : float option
        GirderPanels : GirderPanels
        Remarks : string option
        GeneralNotes : string option
        TcAddedConcentratedLoads : ConcentratedLoad list
        BcAddedConcentratedLoads : ConcentratedLoad list
        TcUniformLoad : float option
        BcUniformLoad : float option
        TcBendCheck : float option
        BcBendCheck : float option
        MainDriftLoad : UniformLoad list  
        FixedMomentLeft : float option
        FixedMomentRight : float option
        LateralMoment : LateralMoment option
        AxialLoadLeft : AxialLoad option
        AxialLoadRight  : AxialLoad option
        MethodOfAxialTransferLeft : MethodOfAxialTransfer option
        MethodOfAxialTransferRight : MethodOfAxialTransfer option
    }


type Joist = 
    {
        Mark : string
        Quantity: int
        Depth: int
        DesignationType: string
        Loading: string
        OverallLengthFt : FtIn
        Tcxl : Tcx option
        Tcxr : Tcx option
        Bcxl : Bcx option
        Bcxr : Bcx option
        BrgDepthLeft : float
        BrgDepthRight : float
        SlopeInfo : SlopeInfo
        LeSeatSlots : SeatSlot option
        ReSeatSlots : SeatSlot option
        NetUplift : string
        AdLoad : float option
        Remarks : string option
        GeneralNotes : string option
        TcAddedConcentratedLoads : ConcentratedLoad list
        BcAddedConcentratedLoads : ConcentratedLoad list
        TcUniformLoad : float option
        BcUniformLoad : float option
        TcBendCheck : float option
        BcBendCheck : float option
        MainDriftLoad : UniformLoad list  
        FixedMomentLeft : float option
        FixedMomentRight : float option
        LateralMoment : LateralMoment option
        AxialLoadLeft : AxialLoad option
        AxialLoadRight  : AxialLoad option
        MethodOfAxialTransferLeft : MethodOfAxialTransfer option
        MethodOfAxialTransferRight : MethodOfAxialTransfer option
    }

type AdditionalNetUpliftLoads = (string*UniformLoad list) list






