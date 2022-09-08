module DESign.BomTools.Vulcraft.Import.Dtos

open DESign.BomTools
open DESign.Helpers

type Joist = 


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

      FixedMoment : {| Left : string option ; Right : string option|}
      LateralMoment : {| WorE : string option ; Left : string option ; Right : string option |}
      IReq : string option
      LeftEndTopChordAxial : {| W : string option ; E : string option ; Em : string option|}
      RightEndTopChordAxial : {| W : string option ; E : string option ; Em : string option|}
  }

type Girder = 


  {
      Mark : string option
      Quantity : string option
      DesignationInfo : {| Depth : string option ; DesignationType : string option ; N : string option ; Load : string option |}
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
      LeTcHoles :
        {|
            Size : string option
            Location : {| Feet : string option ; Inch : string option |}
            DistanceBetweenHoles : string option
            Gage : string option
        |}
      ReTcHoles :
        {|
            Size : string option
            Location : {| Feet : string option ; Inch : string option |}
            DistanceBetweenHoles : string option
            Gage : string option
        |}
      NetUplift : string option
      AdLoad : string option
      GirderGeometry:
        {|
            A : {|Feet : string option ; Inch : string option|}
            N : string option
            Space : {|Feet : string option ; Inch : string option|}
            B : {|Feet : string option ; Inch : string option|}
        |}
      Remarks : string option
      GeneralRemarks : string option
      TcConcentratedLoads : {| Load : string option ; DistanceFt : string option ; DistanceIn : string option|} list
      BcConcentratedLoads : {| Load : string option ; DistanceFt : string option ; DistanceIn : string option|} list
      TcUniformLoad : string option
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

      FixedMoment : {| Left : string option ; Right : string option|}
      LateralMoment : {| WorE : string option ; Left : string option ; Right : string option |}
      IReq : string option
      LeftEndTopChordAxial : {| W : string option ; E : string option ; Em : string option|}
      RightEndTopChordAxial : {| W : string option ; E : string option ; Em : string option|}
  }






