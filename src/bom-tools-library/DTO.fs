module DESign.BomTools.Dto

type NoteDto =
    { ID : string option
      Notes : string option seq }
      
type LoadDto =
    { ID : string option
      Type : string option
      Category : string option
      Position : string option
      Load1Value : float option
      Load1DistanceFt : float option
      Load1DistanceIn : float option
      Load2Value : float option
      Load2DistanceFt : float option
      Load2DistanceIn : float option
      Ref : string option
      LoadCases : string option
      Remarks : string option }
 
type JoistDto =
    { Mark : string option
      Quantity : int option
      JoistSize : string option
      OverallLengthFt : float option
      OverallLengthIn : float option
      TcxlLengthFt : float option
      TcxlLengthIn : float option
      TcxlType : string option
      TcxrLengthFt : float option
      TcxrLengthIn : float option
      TcxrType : string option
      SeatDepthLeft : float option
      SeatDepthRight : float option
      BcxlLengthFt : float option
      BcxlLengthIn : float option
      BcxlType : string option
      BcxrLengthFt : float option
      BcxrLengthIn : float option
      BcxrType : string option
      PunchedSeatsLeftFt : float option
      PunchedSeatsLeftIn : float option
      PunchedSeatsRightFt : float option
      PunchedSeatsRightIn : float option
      PunchedSeatsGa : float option
      OverallSlope : float option
      Notes : string option }

type AdditionalJoistOnGirderDto =
    { LocationFt : string option
      LocationIn : string option
      LoadValue : float option }

type GirderExcessInfoLine =
  { Mark : string option
    NearFarBoth : string option
    AFt : float option
    AIn : float Option
    PanelQuantity : int option
    PanelLengthFt : float option
    PanelLengthIn : float option
    BFt : float option
    BIn : float option
    Load : float option
    AdditionalJoistLoads : AdditionalJoistOnGirderDto seq
}

type GirderDto =
    { Mark : string option
      Quantity : int option
      GirderSize : string option
      OverallLengthFt : float option
      OverallLengthIn : float option
      TcWidth : float option
      TcxlLengthFt : float option
      TcxlLengthIn : float option
      TcxlType : string option
      TcxrLengthFt : float option
      TcxrLengthIn : float option
      TcxrType : string option
      SeatDepthLeft : float option
      SeatDepthRight : float option
      BcxlLengthFt : float option
      BcxlLengthIn : float option
      BcxrLengthFt : float option
      BcxrLengthIn : float option
      PunchedSeatsLeftFt : float option
      PunchedSeatsLeftIn : float option
      PunchedSeatsRightFt : float option
      PunchedSeatsRightIn : float option
      PunchedSeatsGa : float option
      OverallSlope : float option
      Notes : string option
      NumKbRequired : int option
      ExcessInfo : GirderExcessInfoLine seq
      PanelLocations : float list option
      AdditionalJoistLoads : AdditionalJoistOnGirderDto seq }




