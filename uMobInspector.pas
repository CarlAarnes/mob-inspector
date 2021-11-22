(*******************************************************************************
                              Company AS
                              Copyright 2013 - 2014

Description: A clone of the Mob-inspector from Some Company.

                              Mob-Inspector clone v1
                              Developer: Carl Aarnes

Using information from documents:

\\SomeServer\swbe004B(Old G).doc
\\SomeServer\MobII_Communication_User_1 1(NG).pdf

*******************************************************************************)

unit uMobInspector;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, CPort, StdCtrls, DateUtils, CPortCtl, Spin, ComCtrls, SetupAPI, HID,
  SynEdit, SynMemo, cxControls, cxContainer, cxEdit, cxTextEdit, cxMaskEdit,
  cxDropDownEdit, cxCalendar, SynEditHighlighter, SynHighlighterGeneral, Menus;

type ByteArray = array [0..24] of byte;
     TByteArr = array of byte;


{Enums for ABPM commands}
type
  TABPMCommand = ( ReadNULL = 0, ReadNumberOfMeasurements = 102, ReadAllNew = 103, ProgramPatientKey = 121,
    ReadPatientKey = 122, ReadSerialNumber = 124, ReadTime = 127, ManoMeterMode = 136, EndManoMeterMode = 137,
    SetTime = 141, ReadCalibrationTime = 143, TestDisplay = 150, ReadBatteryVoltage = 151, TestEEPROM = 152,
    TestButtons = 154, ReadMemoVoltage = 155, TestBuzzer = 156, ReadEEPROM = 157, ReadVersion = 160, GetProtocol = 167, SetProtocol = 168,
    DeleteHeader = 185, ProgramUserData = 200, ReadUserData = 201, EraseUserData = 202, ReadMeasurementShortError = 205,
    ReadMeasurement = 206, ProgramStandardValues = 220, SelectStandardValues = 221, EraseAllMeasurements = 225,
    WriteHeader = 230, ReadHeader = 231 );

{Enums for old ABPM commands}
type
  TABPMOldCommand = (ReadTimeOld = 27, ManoMeterModeOld = 36, SetTimeOld = 41, ReadCalibrationTimeOld = 43, ReadMemVoltOld = 55, DeleteHeaderOld = 85,
  ReadAllOld = 101);

{Enums for ABPM response}
//these values are responses returned from ABPM device.
type
  TABPMResponse = (NoMeasurements = 7, PatientKey = 22, SerialNumber = 24, BatteryVoltage = 53, MemoVoltage = 55, Version = 61,
    Protocol = 67, Erased = 70, NumberOfMeasurements = 103, FailedRead = 104, UserData = 203, NotAvailable = 207,
    StandardValues = 223, HeaderRead = 232 );

type
  TMeasurement = class(TObject)
  private
    FSys        :byte;
    FDia        :byte;
    FMeanVal    :byte;
    FPulse      :byte;
    FHour       :integer;
    FMin        :integer;
    FDay        :integer;
    FMonth      :integer;
    FYear       :integer;
    FError      :byte;
    FBatVolt    :integer;
  public
    constructor Create();
    destructor Destroy(); override;
    property Sys:byte read FSys;
    property Dia:byte read FDia;
    property Mean:byte read FMeanVal;
    property Pulse:byte read FPulse;
    property Hour:integer read FHour;
    property Min:integer read FMin;
    property Day:integer read FDay;
    property Month:integer read FMonth;
    property Year:integer read FYear;
    property Error:byte read FError;
    property Volt:integer read FBatVolt;
  end;

type
  Tmåling = record
    Sys       :Byte;
    Dia       :Byte;
    Puls      :Byte;
    Mean      :Byte;
    Dato      :TDateTime;
    Voltage   :Byte;
    ErrorCode :Byte;
  end;

type
  TLastErrors = record
    Dato      :TDateTime;
    ErrorCode :Byte;
  end;

type
  PDevBroadcastHdr = ^TDevBroadCastHdr;
  TDevBroadCastHdr = packed record
    dbcd_size         :DWord;
    dbcd_devicetype   :DWord;
    dbcd_reserved     :DWord;
  end;

type
  PDevBroadcastDeviceInterface = ^DEV_BROADCAST_DEVICEINTERFACE;
  DEV_BROADCAST_DEVICEINTERFACE = record
    dbcc_size         :DWord;
    dbcc_devicetype   :DWord;
    dbcc_reserved     :DWord;
    dbcc_classguid    :TGuid;
    dbcc_name         :Char;
  end;
  TDevBroadcastDeviceInterface = DEV_BROADCAST_DEVICEINTERFACE;

  const
    DBT_DEVICEARRIVAL         = $8000;
    DBT_DEVICEREMOVECOMPLETE  = $8004;
    DBT_DEVTYPE_DEVICEINTERFACE = $00000005;
    DBT_DEVICEQUERYREMOVE = $8001;
    DBT_DEVTYP_PORT = 3;

type
  TDeviceNotifyProc = procedure(Sender: TObject; const Devicename:String) of Object;
  TDeviceNotifier = class
  private
    hRecipient: hWnd;
    FNotificationHandle: Pointer;
    FDeviceArrival: TDeviceNotifyProc;
    FDeviceRemoval: TDeviceNotifyProc;
    procedure WndProc(var Msg: TMessage);
  public
    constructor Create(GUID_DEVINTERFACE: TGuid);
    property OnDeviceArrival: TDEviceNotifyProc read FDeviceArrival write FDeviceArrival;
    property OnDeviceRemoval: TDeviceNotifyProc read FDeviceRemoval write FDeviceRemoval;
    destructor Destroy; Override;
  end;

type
  TForm1 = class(TForm)
    btnOpenComport: TButton;
    btnCloseComport: TButton;
    btnBattVolt: TButton;
    btnMemVolt: TButton;
    btnSetTime: TButton;
    btnSetCalibTime: TButton;
    ComLed1: TComLed;
    ComComboBox1: TComComboBox;
    btnManoMode: TButton;
    btnReadNoMeas: TButton;
    btnEraseUserData: TButton;
    btnDeleteMeas: TButton;
    btnEndManoMode: TButton;
    btnSetPatKey: TButton;
    btnSetActiveProtocol: TButton;
    btnReadUserData: TButton;
    btnProgramUserData: TButton;
    Edit1: TEdit;
    SpinEdit1: TSpinEdit;
    DateTimePicker1: TDateTimePicker;
    SpinEdit2: TSpinEdit;
    Edit2: TEdit;
    SpinEdit3: TSpinEdit;
    btnClrScreen: TButton;
    btnReadHeader: TButton;
    btnReadDefaultVal: TButton;
    btnProgramDefValues: TButton;
    SpinEdit4: TSpinEdit;
    SpinEdit5: TSpinEdit;
    SpinEdit6: TSpinEdit;
    SpinEdit7: TSpinEdit;
    btnReadMeasurements: TButton;
    SpinEdit8: TSpinEdit;
    btnDelHeader: TButton;
    btnSynktest: TButton;
    SpinEdit9: TSpinEdit;
    ComPort1: TComPort;
    btnBuzzerTest: TButton;
    btnButtonTest: TButton;
    btnDisplayTest: TButton;
    btnReadSerialNo: TButton;
    btnGetVersionNo: TButton;
    btnReadTime: TButton;
    btnReadCalibTime: TButton;
    btnReadActiveProtocol: TButton;
    btnReadPatKey: TButton;
    btnTestEeprom: TButton;
    btnReadMem: TButton;
    ComComboBox2: TComComboBox;
    btnEnum: TButton;
    btnWriteHeader: TButton;
    SynMemo1: TSynMemo;
    CheckBox1: TCheckBox;
    DateTimePicker2: TDateTimePicker;
    CheckBox2: TCheckBox;
    DateTimePicker3: TDateTimePicker;
    SynGeneralSyn1: TSynGeneralSyn;
    btnSearchMeasurements: TButton;
    PopupMenu1: TPopupMenu;
    N11: TMenuItem;
    edtHeader: TEdit;
    Label1: TLabel;
    Edit3: TEdit;
    procedure btnOpenComportClick(Sender: TObject);
    procedure btnCloseComportClick(Sender: TObject);
    procedure btnBuzzerTestClick(Sender: TObject);
    procedure btnDisplayTestClick(Sender: TObject);
    procedure lesMinnespenning(Sender: TObject);
    procedure lesBatterispenning(Sender: TObject);
    procedure lesSerieno(Sender: TObject);
    procedure btnButtonTestClick(Sender: TObject);
    procedure lesVersjon(Sender: TObject);
    procedure lesTid(Sender: TObject);
    procedure lesCaltid(Sender: TObject);
    procedure settTid(Sender: TObject);
    procedure lesProtokoll(Sender: TObject);
    procedure lesantMeas(Sender: TObject);
    procedure slettBrukerdata(Sender: TObject);
    procedure slettMeas(Sender: TObject);
    procedure manoMode(Sender: TObject);
    procedure stopmanoMode(Sender: TObject);
    procedure lesPatKey(Sender: TObject);
    procedure settPatKey(Sender: TObject);
    procedure programBrukerdata(Sender: TObject);
    procedure ProgrammerStandardValues(Sender: TObject);
    procedure VelgStandardVerdier(Sender: TObject);
    procedure btnSetActiveProtocolClick(Sender: TObject);
    procedure btnReadUserDataClick(Sender: TObject);
    procedure btnClrScreenClick(Sender: TObject);
    procedure btnReadHeaderClick(Sender: TObject);
    procedure btnReadMeasurementsClick(Sender: TObject);
    procedure btnDelHeaderClick(Sender: TObject);
    procedure btnSynktestClick(Sender: TObject);
    procedure btnTestEepromClick(Sender: TObject);
    procedure btnReadEeprom(Sender: TObject);
    procedure btnUsbEnum(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure SynMemo1SpecialLineColors(Sender: TObject; Line: Integer;
      var Special: Boolean; var FG, BG: TColor);
    procedure SynMemo1Click(Sender: TObject);
    procedure btnWriteHeaderClick(Sender: TObject);
    procedure N11Click(Sender: TObject);
    procedure btnSearchMeasurementsClick(Sender: TObject);
  private
    { Private declarations }
    DeviceNotifier: TDeviceNotifier;
    function CalcChecksum(data: array of byte):integer;
    function Sync():boolean;
    function ABPMCommand(Command:TABPMCommand):boolean; Overload;
    function ABPMCommand(Command:TABPMCommand; out Data:ByteArray):boolean; Overload;
    function ABPMCommand(Command:TABPMCommand; aData:ByteArray; out Data:ByteArray):boolean; Overload;

    function ABPMCommand(Command:TABPMOldCommand; out Data:ByteArray):boolean; Overload;
    function ABPMCommand(Command:TABPMOldCommand):boolean; overload;
    function ABPMCommand(Command:TABPMOldCommand; aData: ByteArray; out Data: ByteArray):boolean; Overload;
    function StrToByte(const Value:String):TByteArr;

  public
    { Public declarations }
    procedure arrival(Sender:TObject; const DeviceName:string);
    procedure removal(Sender:TObject; const Devicename:string);
  end;

var
  Form1 :TForm1;
  GUID_DEVINTERFACE_USB_DEVICE :TGuid = '{A5DCBF10-6530-11D2-901F-00C04FB951ED}';
  ShadowStrList: TStringlist;
  glStartpos: integer;
  glEndpos: integer;
  glLength: integer;
implementation

{$R *.dfm}

uses uBat;

constructor TDeviceNotifier.Create(GUID_DEVINTERFACE: TGUID);
var
  NotificationFilter: TDevBroadcastDeviceInterface;
begin
  inherited Create;
  hRecipient := Classes.AllocateHwnd(WndProc);
  ZeroMemory(@NotificationFilter,SizeOf(NotificationFilter));
  NotificationFilter.dbcc_size := SizeOf(NotificationFilter);
  NotificationFilter.dbcc_devicetype := DBT_DEVTYPE_DEVICEINTERFACE;
  NotificationFilter.dbcc_reserved := 0;
  NotificationFilter.dbcc_classguid := GUID_DEVINTERFACE_USB_DEVICE;
  FNotificationHandle := RegisterDeviceNotification(hRecipient,@NotificationFilter,DEVICE_NOTIFY_WINDOW_HANDLE);
end;



procedure TDeviceNotifier.WndProc(var Msg: TMessage);
var
  Dbi: PDevBroadcastDeviceInterface;
begin
  with Msg do
    if (Msg = WM_DEVICECHANGE) and ((WParam = DBT_DEVICEARRIVAL) or (WParam = DBT_DEVICEREMOVECOMPLETE))
    then
    try
      Dbi := PDevBroadcastDeviceInterface(LParam);
      if Dbi.dbcc_devicetype = DBT_DEVTYPE_DEVICEINTERFACE then
      begin
        if WParam = DBT_DEVICEARRIVAL then
        begin
          if Assigned(FDeviceArrival) then
            //FDeviceArrival(Self,PChar(@Dbi.dbcc_name));
            OnDeviceArrival(Self,PChar(addr(Dbi.dbcc_name)));
          end
          else
          if WParam = DBT_DEVICEREMOVECOMPLETE then
          begin
            if Assigned(FDeviceRemoval) then
              FDeviceRemoval(Self,PChar(@Dbi.dbcc_name));
          end;
      end;

    except
      Result := DefWindowProc(hRecipient,Msg, WParam, LParam);
    end
    else
      Result := DefWindowProc(hRecipient,Msg,WParam,LParam);
end;

destructor TDeviceNotifier.Destroy;
begin
  UnregisterDeviceNotification(FNotificationHandle);
  Classes.DeallocateHwnd(hRecipient);
  inherited;
end;


constructor TMeasurement.Create();
begin
  inherited;
end;

destructor TMeasurement.Destroy;
begin
  inherited;
end;

procedure PrintErrors(var Last100Errors: array of TLastErrors);
var
  Errorlist: TStringList;
  antall  :integer;
  i       :smallint;
  ErrorNo :smallint;
  ErrorStr: String;
begin
  Errorlist := TStringList.Create();
  try
    Errorlist.LoadFromFile('Feilkode.spr');
    ErrorNo := 1;
    antall := length(Last100Errors);
    Form1.SynMemo1.Lines.Add(#13);
    Form1.SynMemo1.Lines.Add('De siste 100 feilmeldinger lagret i apparatet.');
    Form1.Synmemo1.Lines.Add('Nr        Tidspunkt            Error');
    For i := 0 to antall-1 do
    begin
      if (int(Last100Errors[i].ErrorCode) > 0) then
        ErrorStr := ErrorList.Strings[Last100Errors[i].ErrorCode]
        //Errorlist.
      else
        ErrorStr := '';

      Form1.Synmemo1.Lines.Add(IntToStr(ErrorNo) + ':   ' +  DateTimeToStr(Last100Errors[i].Dato) + '  ' + IntToStr(Last100Errors[i].ErrorCode) + ' - ' + ErrorStr);
      inc(ErrorNo);
    end;
  finally
    Errorlist.Free();
  end;
end;

function SecondsToDateTime(datobyte1,datobyte2,datobyte3,datobyte4:Byte):TDateTime;
var
  Startdate     :TDateTime;
  EndDate       :TDateTime;
  Newdate       :TDateTime;
  days,hrs,mins :cardinal;
  sec           :int64;
  sumOfSeconds  :cardinal;
begin
  StartDate := EncodeDateTime(2005,1,1,0,0,0,0);

// Dato fra APBM minne lagret som 4 bytes!
// konverter bytes til antall sekunder siden 1.1.2005
  sumOfSeconds := datobyte4 shl 24 + datobyte3 shl 16 + datobyte2 shl 8 + datobyte1;
  sec := sumOfSeconds;

  days := Trunc(sec/secsPerDay);                                                      //sec->min->Hrs->Days
  hrs := Trunc((sec/(secsPerMin*MinsPerHour))  - (days*HoursPerDay));                 //sec->min->Hrs - days*24
  mins := Trunc((sec/secsPerMin) - (days*HoursPerday*MinsPerHour + Hrs*MinsPerHour)); //minutes left                                                    //sec->min - Hrs*60
  //sec := integer(sec) mod 60;                                                         //mod only takes integer !!!

  Newdate := IncDay(StartDate,days);
  NewDate := IncHour(NewDate,hrs);
  NewDate := IncMinute(NewDate,mins);
  //NewDate := IncSecond(NewDate,sec);

  result :=  NewDate;
end;

procedure Tform1.arrival(Sender: TObject; const DeviceName: string);
begin
  SynMemo1.Lines.Add('Ny usb enhet funnet: ' + DeviceName);
  EnumComports(ComComboBox1.Items);
end;

procedure TForm1.removal(Sender: TObject; const Devicename: string);
begin
  SynMemo1.Lines.Add('Usb enhet fjernet');
  EnumComports(ComComboBox1.Items);
end;

procedure Tform1.ProgrammerStandardValues(Sender: TObject);
var
  WriteBuffer     :ByteArray;
  ReadBuffer      :ByteArray;
  diastolicDay    :byte;
  systolicDay     :byte;
  diastolicNight  :byte;
  systolicNight   :byte;
begin
  FillChar(WriteBuffer,sizeOf(WriteBuffer),0);

  WriteBuffer[0] := spinedit6.Value;                                                          // Sett standardverdier for blodtrykk.
  WriteBuffer[1] := spinedit4.Value - 45;
  WriteBuffer[2] := spinedit7.Value;
  WriteBuffer[3] := spinedit5.Value - 45;

  ABPMCommand(ProgramStandardValues,WriteBuffer,ReadBuffer);
end;


procedure TForm1.VelgStandardVerdier(Sender: TObject);
var
  ReadBuffer      :ByteArray;
  diastolicDay    :byte;
  systolicDay     :byte;
  diastolicNight  :byte;
  systolicNight   :byte;
begin
  ABPMCommand(SelectStandardValues,ReadBuffer);
  if ReadBuffer[2] = byte(StandardValues) then
  begin
    diastolicDay := ReadBuffer[3];
    systolicDay := ReadBuffer[4] + 45;
    diastolicNight := ReadBuffer[5];
    systolicNight := ReadBuffer[6] + 45;
  end;
  //litt typecasting, men funker....
  Synmemo1.Lines.Add('Dag: ' + intToStr(integer(systolicDay)) + '/' + intToStr(integer(diastolicDay)));
  Synmemo1.Lines.Add('Natt: ' + IntToStr(integer(systolicNight)) + '/' + intToStr(integer(diastolicNight)));
end;

function Tform1.StrToByte(const Value: String):TByteArr;
var
  i  :integer;
begin
  SetLength(Result, Length(Value));
  for i := 0 to Length(Value)- 1 do
    Result[i] := Ord(Value[i+1]);
end;

procedure Tform1.programBrukerdata(Sender: TObject);
var
  WriteBuffer  :ByteArray;
  ReadBuffer   :ByteArray;
  blokk        :Integer;
  tmpArray     :TByteArr;
  i            :integer;
begin
  FillChar(WriteBuffer,SizeOf(WriteBuffer),0);
  blokk := SpinEdit2.Value;
  WriteBuffer[0] := blokk;

  tmpArray := StrToByte(Edit2.Text);

  for I := 0 to length(tmpArray)-1 do
    WriteBuffer[i+1] := tmpArray[i];
  //ABPMCommand(200,inbuff,utbuff);
  ABPMCommand(ProgramUserData,WriteBuffer,ReadBuffer);
end;

procedure Tform1.settPatKey(Sender: TObject);
var
  PatKey      :string;
  WriteBuffer :ByteArray;
  ReadBuffer  :ByteArray;
begin
  //PatKey := '1234-5678';
  PatKey := Edit1.Text;
  Move(PatKey[1], WriteBuffer, Length(PatKey) * SizeOf(Char));
  //ABPMCommand(121,inbuff,utbuff);
  ABPMCommand(ProgramPatientKey,WriteBuffer,ReadBuffer);
end;


procedure Tform1.lesPatKey(Sender: TObject);
var
  ReadBuffer  :ByteArray;
  PatKey      :string;
  i           :integer;
begin
  //ABPMCommand(122,buff);
  ABPMCommand(ReadPatientKey,ReadBuffer);
  if ReadBuffer[2] = byte(PatientKey) then
    for i := 3 to 23 do
    begin
      PatKey := PatKey + chr(ReadBuffer[I]);
    end;
  Synmemo1.Lines.Add('Pasient nøkkel: ' + PatKey);
end;


procedure Tform1.manoMode(Sender: TObject);
begin
  //ABPMCommand(136);
  ABPMCommand(ManometerMode);
end;


procedure TForm1.N11Click(Sender: TObject);
var
  Exporttext: String;
  foundpos: integer;
  strList :TStringlist;
  FirstSelectedLine: string;
  LastSelectedLine: string;
  lastpos: integer;
  foundstr: string;
  shadowstrIndex : integer;
  firstIdx,lastIdx: integer;
  I: integer;
begin
  strList := TStringlist.Create();

  FirstSelectedLine := SynMemo1.LineText;
  foundPos := Ansipos(':',FirstSelectedLine);
  foundstr := Copy(FirstSelectedLine,0,foundpos-1);
  firstIdx := StrToInt(foundstr);

  //strList.Add(ShadowStrList.Strings[shadowstrIndex-1]);

  SynMemo1.CaretY := SynMemo1.CharIndexToRowCol(glLength).Line+firstIdx;

  LastSelectedLine := SynMemo1.LineText;
  lastpos := Ansipos(':',LastSelectedLine);
  foundstr := Copy(LastSelectedLine,0,lastpos-1);

  lastIdx := StrToInt(foundstr);

  strList.Add('Systolisk;Diastolisk;Middelarterietrykk;Puls;Tid;Spenning;Feilkode');
  for I := firstIdx-1 to lastIdx-1 do
    strList.Add(ShadowStrList.Strings[I]);
  strList.SaveToFile(Edit3.Text);

  strList.Free();
end;

procedure Tform1.stopmanoMode(Sender: TObject);
begin
  //ABPMCommand(137);
  ABPMCommand(EndManoMeterMode);
end;


procedure Tform1.slettMeas(Sender: TObject);
var
  ReadBuffer  :ByteArray;
begin
  //ABPMCommand(225);
  ABPMCommand(EraseAllMeasurements);
    if ReadBuffer[2] = byte(Erased) then
  SynMemo1.Lines.Add('Målinger slettet!');
end;


procedure Tform1.slettBrukerdata(Sender: TObject);
begin
  //ABPMCommand(202);
  ABPMCommand(EraseUserData);
end;


procedure Tform1.lesantMeas(Sender: TObject);
var
  ReadBuffer  :ByteArray;
  noMeas      :integer;
begin
  //ABPMCommand(102,buff);
  ABPMCommand(ReadNumberOfMeasurements,ReadBuffer);
  if (ReadBuffer[2] = byte(NumberOfMeasurements)) then
    noMeas := ReadBuffer[3]*256 + ReadBuffer[4];
  Synmemo1.Lines.Add('Antall målinger: ' + IntToStr(noMeas));
end;


procedure Tform1.lesProtokoll(Sender: TObject);
var
  ReadBuffer  :ByteArray;
  protokoll   :integer;
begin
  //ABPMCommand(167,buff);
  ABPMCommand(GetProtocol,ReadBuffer);
  if ReadBuffer[2] = byte(Protocol) then
    protokoll := ReadBuffer[3];
  SynMemo1.Lines.Add('Aktiv protokoll: ' + IntToStr(protokoll));
end;


procedure TForm1.lesTid(Sender: TObject);
var
  ReadBuffer    :ByteArray;
  nytid         :TDateTime;
  tid           :string;
  aar,mnd,dag   :integer;
  time,min      :integer;
begin
  //ABPMCommand(27,buff);
  ABPMCommand(ReadTimeOld,ReadBuffer);

  if ReadBuffer[2] = 27 then
  begin
    aar := ReadBuffer[3] + 2000;
    mnd := ReadBuffer[4];
    dag := ReadBuffer[5];
    time := ReadBuffer[6];
    min := ReadBuffer[7];                                                           // ingen sek. fra apparat!
  end;

  try
    nytid := EncodeDateTime(aar,mnd,dag,time,min,0,0);
    tid := DateTimeToStr(nytid);
    //tid := IntToStr(Dag)+ '/' + IntToStr(mnd) + '/' + IntToStr(aar) + '  ' + IntToStr(time) + ':' + IntToStr(min);
    SynMemo1.Lines.Add('Klokkeslett i apparatet: ' + tid);
  Except
    On E: EConvertError do
    ShowMessage('tulledato');
  end;
end;


procedure TForm1.settTid(Sender: TObject);
var
  WriteBuffer       :ByteArray;
  tmpBuffer         :ByteArray;
  aar,mnd,dag       :word;
  time,min,sek,ms   :word;
  tid               :TDateTime;
begin
  tid := now;
  DecodeDateTime(tid,aar,mnd,dag,time,min,sek,ms);

  WriteBuffer[0] := time;
  WriteBuffer[1] := min;
  WriteBuffer[2] := sek;

  WriteBuffer[3] := dag;
  WriteBuffer[4] := mnd;
  WriteBuffer[5] := aar - 2000;;

  //ABPMCommand(141,buff,tmpbuff);
  ABPMCommand(SetTime,WriteBuffer,tmpBuffer);
end;


procedure TForm1.lesCaltid(Sender: TObject);
{Leser tidspunkt for siste kalibrering.}
var
  ReadBuffer    :ByteArray;
  nytid         :TDateTime;
  dato          :string;
  tid           :string;
  aar,mnd,dag   :integer;
  time,min      :integer;
begin
  //ABPMCommand(143,buff);
  ABPMCommand(ReadCalibrationTime,ReadBuffer);
  if ReadBuffer[2] = 43 then
  begin
    aar := ReadBuffer[3] + 2000;
    mnd := ReadBuffer[4];
    dag := ReadBuffer[5];
    time := ReadBuffer[6];
    min := ReadBuffer[7];                                                       // ingen sek. fra apparat!
  end;

  LongTimeFormat := 'hh:mm';

  try
    nytid := EncodeDateTime(aar,mnd,dag,time,min,0,0);
    dato := DateToStr(nytid);
    tid := TimeToStr(nytid);
    //tid := IntToStr(Dag)+ '/' + IntToStr(mnd) + '/' + IntToStr(aar) + '  ' + IntToStr(time) + ':' + IntToStr(min);
    SynMemo1.Lines.Add('Siste kalibrering: ' + dato + ' ' + tid);
  Except
    On E: EConvertError do
    begin
      ShowMessage('datofeil');
      raise;
    end;
  end;
end;

procedure TForm1.btnOpenComportClick(Sender: TObject);
begin
  //Comport1.Port := ComComboBox1.ComPort.Port;
  Comport1.Open;
end;

procedure TForm1.btnReadUserDataClick(Sender: TObject);
{Brukerminnet er delt opp i 10 blokker på 20 bytes. (blokk 0-9)
Program bruker blokk 0 for å lagre informasjon om CuffSize.}
var
  blokk       :integer;
  ReadBuffer  :ByteArray;
  WriteBuffer :ByteArray;
  i           :integer;
  UserString  :string;
begin
  FillChar(WriteBuffer,sizeOf(WriteBuffer),0);

  blokk := spinedit3.Value;
  WriteBuffer[0] := blokk;

  ABPMCommand(ReadUserData,WriteBuffer,ReadBuffer);

  if ReadBuffer[2] = byte(UserData) then
  begin
    SynMemo1.Lines.Add('Leser data fra blokk: ' + intToStr(blokk));
    for I := 4 to sizeOf(ReadBuffer)-1 do
      UserString := UserString + chr(ReadBuffer[i]);
    SynMemo1.Lines.Add(UserString);
  end;
end;

procedure TForm1.btnSearchMeasurementsClick(Sender: TObject);
begin
  btnReadEeprom(self);  
end;

procedure TForm1.btnSetActiveProtocolClick(Sender: TObject);
var
  ReadBuffer      :ByteArray;
  WriteBuffer     :ByteArray;
begin
  WriteBuffer[0] := SpinEdit1.Value;
  ABPMCommand(SetProtocol,WriteBuffer,ReadBuffer);
end;

procedure TForm1.btnClrScreenClick(Sender: TObject);
begin
  SynMemo1.Lines.Clear;
end;

procedure TForm1.btnWriteHeaderClick(Sender: TObject);
var
  ReadBuffer    :ByteArray;
  WriteBuffer   :ByteArray;
  HeaderTxt     :string;
  blokk         :integer;
  I             :integer;
  Test          :ByteArray;
begin
  Fillchar(WriteBuffer,sizeOf(WriteBuffer),0);
  blokk := 0;
  WriteBuffer[0] := blokk;

  //Ser ut til at CRC ikke er påkrevd her....
  //CRC er summen av 70 etterfølgende bytes av CRC.
  //WriteBuffer[1] := HiByte(0);                                                  //Highbyte CRC;
  //WriteBuffer[2] := LoByte(0);                                                  // Lowbyte CRC

  HeaderTxt := edtHeader.Text;

  for i := 0 to length(HeaderTxt) do
  begin
    WriteBuffer[i+2] := Byte(HeaderTxt[i]);
  end;
  ABPMCommand(WriteHeader,WriteBuffer,ReadBuffer);
end;

function checkandfill(argValue:integer):string;
begin
  if ArgValue < 10 then
    result := '00' +IntToStr(argValue)
  else
  if (ArgValue >= 10) and (argValue <100) then
    result := '0' + IntToStr(argValue)
  else
    result := IntToStr(argValue);
end;

procedure TForm1.btnReadEeprom(Sender: TObject);
var
  ReadBuffer      :ByteArray;
  WriteBuffer     :ByteArray;
  EEPROMBuffer    :Array [0..8192] of byte;
  i               :integer;
  line            :string;
  løpenr          :integer;
  fill            :string;
  checknr         :integer;
  ABPMDato        :TDateTime;
  ABPMBaseAdr     :integer;
  MemSys          :Byte;
  MemDia          :Byte;
  MemMean         :Byte;
  MemHR           :Byte;
  memBattVLo      :smallint;
  memBattVHi      :Byte;
  MemErr          :Byte;
  memBatt         :Integer;
  Last600         :Array of Tmåling;
  Last100Errors   :Array of TLastErrors;
  Minnemålepos    :Integer;
  Last100Adr      :Integer;
  EndAdr          :Integer;
  j               :Integer;
  errorpos        :Integer;
  BuffAdr         :Integer;
  TestDate        :Tdate;
  FraDato         :Tdate;
  TilDato         :Tdate;
  FunnDato        :Tdate;
  FunnetMålinger  :Array of Tmåling;
  r               :integer;
  s               :integer;
  nFound          :smallint;

  procedure PrintMålinger(var Målinger: array of Tmåling);
  var
    antall      :integer;
    i           :smallint;
    MeasNumber  :Smallint;
    fyll        :string;
  begin
    MeasNumber := 1;
    antall := Length(Målinger);
    ShadowStrList.Clear();

    {if Checkbox2.Checked = true then
      glStartpos := SynMemo1.CaretY;
    SynMemo1.OverwriteCaret := false; }
    For i:= 0 to antall-1 do
    begin
      if MeasNumber < 10 then fyll := '00'
        else
      fyll := '0';
      Synmemo1.Lines.Add(fyll + IntToStr(MeasNumber) + ':    ' + Checkandfill(Målinger[i].Sys) + '   ' + Checkandfill(Målinger[i].Dia) + '   ' + Checkandfill(Målinger[i].Mean) + '   ' + Checkandfill(Målinger[i].Puls)
      + '          ' + DateTimeToStr(Målinger[i].Dato) + '     ' + Checkandfill(Målinger[i].Voltage) + '     ' + Checkandfill(Målinger[i].ErrorCode));
      SynMemo1.SetFocus();
      if (Checkbox2.Checked = true) then
        ShadowStrList.Add(IntToStr(Målinger[i].Sys) + ';' + IntToStr(Målinger[i].Dia) + ';' + IntToStr(Målinger[i].Mean) + ';' + IntToStr(Målinger[i].Puls) + ';' + DateTimeToStr(Målinger[i].Dato) + ';' + IntToStr(Målinger[i].Voltage) + ';' + IntToStr(Målinger[i].ErrorCode));
      inc(MeasNumber);
    end;

   { SynMemo1.Lines.Add(#13#10);
    
    if Checkbox2.Checked = true then
      glEndpos := SynMemo1.CaretY; }
  end;

begin
  i := 1;
  løpenr := 0;

  FillChar(EEPROMBuffer,sizeOf(EEPROMBuffer),0);
  Comport1.Open();
  try
    FillChar(WriteBuffer,sizeOf(WriteBuffer),0);
    WriteBuffer[0] := 2;
    WriteBuffer[1] := Byte(ReadEEPROM);
    WriteBuffer[23] := CalcChecksum(WriteBuffer);
    WriteBuffer[24] := 3;

    //if Sync() then
      Comport1.Write(WriteBuffer,sizeOf(WriteBuffer));

    Comport1.Read(EEPROMBuffer,SizeOf(EEPROMBuffer));

    if Checkbox2.Checked = false then
    begin
      Synmemo1.Lines.Add(sLineBreak + 'Reading numerical byte values from EEPROM:' + #13);

      while i < sizeOf(EEpromBuffer) do
      begin
        if løpenr < 10 then fill := '000';
        if løpenr > 10 then fill := '00';
        if (løpenr < 1000) and (løpenr > 100) then fill := '0';
        if løpenr > 1000 then fill := '';

        if i >= 8192 then                                                           // ram max 8 kb ?
          break;
        line := fill + IntToStr(løpenr) + ':  ' + checkandfill(EEPROMBuffer[i]) + ' , ' + CheckAndFill(EEPROMBuffer[i+1]) + ' , ' + CheckAndFill(EEPROMBuffer[i+2]) + ' , ' +
        CheckAndFill(EEPROMBuffer[i+3]) + ' , ' + CheckAndFill(EEPROMBuffer[i+4]) + ' , ' + CheckAndFill(EEPROMBuffer[i+5]) + ' , ' +
        CheckAndFill(EEPROMBuffer[i+6]) + ' , ' + CheckAndFill(EEPROMBuffer[i+7]) + ' , ' + CheckAndFill(EEPROMBuffer[i+8]) + ' , ' +
        CheckAndFill(EEPROMBuffer[i+9]) + ' , ' + CheckAndFill(EEPROMBuffer[i+10]) + ' , ' +
        CheckAndFill(EEPROMBuffer[i+11]) + ' , ' + CheckAndFill(EEPROMBuffer[i+12]) + ' , ' + CheckAndFill(EEPROMBuffer[i+13]) + ' , ' +
        CheckAndFill(EEPROMBuffer[i+14]) + ' , ' + CheckAndFill(EEPROMBuffer[i+15]);
        inc(i,16);
        inc(løpenr,16);
        Synmemo1.Lines.Add(line);
      end;
    end;

    ABPMBaseAdr := 811;

    Minnemålepos := 0;
    SetLength(Last600,599);
    if Checkbox2.Checked = false then
    begin
      SynMemo1.Lines.Add(#13);
      Synmemo1.Lines.Add('Nr      Sys   Dia   Map   Hr                   Time            Volt    iErr');
    end;

    While (EEpromBuffer[ABPMBaseAdr+4]<>170) and (EEpromBuffer[ABPMBaseAdr+5]<>170) and (EEpromBuffer[ABPMBaseAdr+6]<>170) and (EEpromBuffer[ABPMBaseAdr+7]<>170) and (minnemålepos<600)  do
    begin
      if Last600[MinneMålepos].Sys <> 0 then
        Last600[MinneMålepos].Sys := EepromBuffer[ABPMBaseAdr]+45
      else
        Last600[MinneMålepos].Sys := EepromBuffer[ABPMBaseAdr];
      Last600[Minnemålepos].Dia :=  EepromBuffer[ABPMBaseAdr+1];
      Last600[Minnemålepos].Puls := EepromBuffer[ABPMBaseAdr+2];
      if Last600[Minnemålepos].Mean <> 0  then
        Last600[Minnemålepos].Mean := EEpromBuffer[ABPMBaseAdr+3]+35
      else
        Last600[Minnemålepos].Mean := EEpromBuffer[ABPMBaseAdr+3];
      Last600[Minnemålepos].Dato := SecondsToDateTime(EEpromBuffer[ABPMBaseAdr+4],EEpromBuffer[ABPMBaseAdr+5],EEpromBuffer[ABPMBaseAdr+6],EEpromBuffer[ABPMBaseAdr+7]);
      Last600[Minnemålepos].ErrorCode := EEpromBuffer[ABPMBaseAdr+8];
      Last600[Minnemålepos].Voltage := EepromBuffer[ABPMBaseAdr+9]*2;
      inc(MinneMålepos);
      inc(ABPMBaseAdr,11);
    end;
    Setlength(Last600,MinneMålepos);

    if Checkbox2.Checked = false then
      PrintMålinger(Last600);

    // Datosøk .....
    if (Checkbox2.Checked = true) then
    begin
      FraDato := DateTimePicker2.Date;
      TilDato := DateTimePicker3.Date;

      SetLength(FunnetMålinger,600);                                            //Max. dyn. array size!

      nFound := 0;
     { for r := 0 to Length(Last600)-1 do
      begin
        case CompareDate(FraDato,Last600[r].Dato) of                            //Er lik FraDato
          0:  begin
                FunnetMålinger[nFound] := Last600[r];
                inc(nFound);
              end;
          1:  ;
         -1:  ;
        end;
      end;  }

      //nFound := 0;
      for s := 0 to Length(Last600)-1 do
      begin
        case CompareDate(TilDato,Last600[s].Dato) of
        0:  begin
              FunnetMålinger[nFound] := Last600[s];
              inc(nFound);
            end;
        1:  begin
              if ((CompareDate(FraDato,Last600[s].Dato) = 0) or (CompareDate(FraDato,Last600[s].Dato) = -1)) then
              begin
                FunnetMålinger[nFound] := Last600[s];
                inc(nFound);
              end;
            end;
       -1:  ;
        end;
      end;

      setLength(FunnetMålinger,nFound);                                         //resize array
      SynMemo1.Lines.Add(#13);
      SynMemo1.Lines.Add('Målinger funnet:');
      Synmemo1.Lines.Add('Nr      Sys   Dia   Map   Hr                   Time            Volt    iErr');
      PrintMålinger(FunnetMålinger);
    end;
    {TestDate := Last600[0].dato;//Målinger[n].Dato;
    for r := 0 to Length(Last600)-1 do
    begin
      case CompareDateTime(TestDate,Last600[r].Dato) of
        0:  ;
        1:  Testdate := Last600[r].Dato;
       -1:  ;
      end;
    end;}

    {SynMemo1.ActiveLineColor := clCream;
    SynMemo1.CaretY := SynMemo1.Lines.Count;
    SynMemo1.SetFocus(); }

// Vis siste 100 feilmeldinger.
    if (CheckBox1.Checked = true) then
    begin
      Last100Adr := 7411;
      BuffAdr := Last100Adr;
      EndAdr := 7914;
      errorpos := 0;

      SetLength(Last100Errors,100);
      For j := 0 to 99 do
      begin
        Last100Errors[errorpos].Dato := SecondsToDateTime(EEpromBuffer[BuffAdr],EEpromBuffer[BuffAdr+1],EEpromBuffer[BuffAdr+2],EEpromBuffer[BuffAdr+3]);
        Last100Errors[errorpos].ErrorCode := EepromBuffer[BuffAdr+4];
        inc(errorpos);
        inc(BuffAdr,5);
      end;
      PrintErrors(Last100Errors);
    end;
    SynMemo1.Lines.Add(#13);
    Synmemo1.CaretY := SynMemo1.Lines.Count;
    Synmemo1.CaretX := 0;
  finally
    Comport1.Close();
  end;
end;

procedure TForm1.btnUsbEnum(Sender: TObject);
const
 USB_CLASS_GUID:TGUID = '{36FC9E60-C465-11CF-8056-444553540000}';
 USB_DEVICE_GUID:TGUID = '{A5DCBF10-6530-11D2-901F-00C04FB951ED}';
 IMI_GUID:TGUID = '{5C668831-3EF4-4B91-B95C-C36FFB9A3C79}';
var
  DevInfo                           :HDEVINFO;
  //DevIntfData                       :SP_DEVICE_INTERFACE_DATA;
  DevInfoDetail                     :PSPDeviceInterfaceDetailData;
  //DevIntfDetailData                 :SP_DEVICE_INTERFACE_DETAIL_DATA_A;
  DevData                           :TSPDevInfoData;
  DevInfoData                       :TSPDeviceInterfaceData;
  dwSize                            :DWord;
  dwMemberIndex                     :DWord;
  reqlen                            :Cardinal;
  res                               :Boolean;
  devPath                           :string;
  err : integer;

  PropertyBuffer : array[0..255] of AnsiChar;
  PropertyRegDataType: DWORD;
  RequiredSize: DWORD;
  propbuffer :string;
begin
  if not LoadSetupAPI() then
    exit;
  DevInfo := SetupDiGetClassDevs(@USB_DEVICE_GUID ,Nil,0,DIGCF_PRESENT or DIGCF_DEVICEINTERFACE);
  if (NativeUInt(DevInfo) = INVALID_HANDLE_VALUE) then
  exit;

  dwMemberIndex := 0;

  repeat
    DevData.cbSize := sizeOf(TSPDevInfodata);
    DevInfoData.cbSize := sizeOf(TSPDeviceInterfaceData);

    res := SetupDiEnumDeviceInterfaces(DevInfo,nil,USB_DEVICE_GUID,dwMemberIndex,DevInfoData);
    if res then
    begin
      //SetupDiGetInterfaceDeviceDetail(DevInfo,@DevInfoData,nil,0,reqlen,nil);  InterfaceDevice ... backward stuff :(
      SetupDiGetDeviceInterfaceDetail(DevInfo,@DevInfoData,nil,0,reqlen,nil);
      dwSize := reqlen;
      Getmem(DevInfoDetail,dwSize);
      DevInfoDetail^.cbSize := SizeOf(TSPDeviceInterfaceDetailData);
      //if SetupDiGetInterfaceDeviceDetail(DevInfo,@DevInfoData,DevInfoDetail,dwSize,reqlen,@DevData) then
      if SetupDiGetDeviceInterfaceDetail(DevInfo,@DevInfoData,DevInfoDetail,dwSize,reqlen,@DevData) then
      begin
        PropertyRegDataType:=0;
        RequiredSize:=0;
        //FillChar(Propertybuffer,10,0);
        if SetupDiGetDeviceRegistryProperty(DevInfo, DevData, SPDRP_FRIENDLYNAME, PropertyRegDataType,  PBYTE(@PropertyBuffer[0]), SizeOf(PropertyBuffer), RequiredSize) then
        Begin
          propbuffer := PChar(@PropertyBuffer);
          SynMemo1.Lines.Add('Friendly Name: ' + Propbuffer);
        End;

        if SetupDiGetDeviceRegistryProperty(DevInfo, DevData, SPDRP_MFG, PropertyRegDataType,  PBYTE(@PropertyBuffer[0]), SizeOf(PropertyBuffer), RequiredSize) then
        Begin
          propbuffer := PChar(@PropertyBuffer);
          SynMemo1.Lines.Add('Produsent: ' + Propbuffer);
        End;

        if SetupDiGetDeviceRegistryProperty(DevInfo, DevData, SPDRP_DEVICEDESC, PropertyRegDataType,  PBYTE(@PropertyBuffer[0]), SizeOf(PropertyBuffer), RequiredSize) then
        Begin
          propbuffer := PChar(@PropertyBuffer);
          SynMemo1.Lines.Add('Beskrivelse: ' + Propbuffer);
        End;

        if SetupDiGetDeviceRegistryProperty(DevInfo, DevData, SPDRP_HARDWAREID, PropertyRegDataType,  PBYTE(@PropertyBuffer[0]), SizeOf(PropertyBuffer), RequiredSize) then
        Begin
          propbuffer := PChar(@PropertyBuffer);
          SynMemo1.Lines.Add('HWID: ' + Propbuffer);
        End;

        DevPath := PChar(@DevInfoDetail.DevicePath);
        Synmemo1.Lines.Add(DevPath);
        SynMemo1.Lines.Add(#13);
      end;
      Freemem(DevInfoDetail);
      inc(dwMemberIndex);
    end;
  until not res;
  //err := Getlasterror();

  //DevIntfData.cbSize := sizeOf(SP_DEVICE_INTERFACE_DATA);
  //SetupDiEnumDeviceInterfaces(DevInfo,Nil,USB_DEVICE_GUID,dwMemberIndex,DevIntfData);
 {   While (GetLastError() <> ERROR_NO_MORE_ITEMS) do
    Begin
      DevData.cbSize := sizeOf(DevData);
      dwSize := 0;
      SetupDiGetDeviceInterfaceDetail(DevInfo,@DevintfData,nil,0,dwSize,nil);
      if dwSize <> 0 then
      begin
        GetMem(DevInfoDetail,dwSize);
        DevInfoDetail.cbSize := SizeOf(TSPDeviceInterfaceDetailData);
        //DevIntfData.cbSize := SizeOf(TSPDeviceInterfaceDetailData);

        SetupDiGetDeviceInterfaceDetail(DevInfo,@DevIntfdata,DevInfoDetail,dwSize,dwSize,@DevData);
      end;
    End; }
  //HidD_GetAttributes();

  SetupDiDestroyDeviceInfoList(DevInfo);
  UnloadSetupAPI();
end;

procedure TForm1.btnSynktestClick(Sender: TObject);
var
  i               :integer;
  n               :integer;
  countOk         :integer;
  countFail       :integer;
  pstOk, pstFail  :integer;
begin
  Comport1.Open();
  try
    Synmemo1.lines.Clear();
    countOk := 0;
    countFail := 0;

    n := spinedit9.Value;

    for i := 1 to n do
    begin
      if Sync() then
      begin
        Synmemo1.Lines.Add(intToStr(i) + '. Sync ok');
        inc(countOk);
      end
      else
      begin
        Synmemo1.Lines.Add(IntToStr(i) + '. Sync failed');
        inc(countFail);
      end;
      Application.ProcessMessages();
    end;
    pstOk := round((countOk/n)*100);
    pstFail := round((countFail/n)*100);

    //memo1.Lines.Add('---------------------------------------------------------------');
    Synmemo1.Lines.Add(StringOfChar('-',58));
    Synmemo1.Lines.Add('Ok: ' + intToStr(countOk) + ', Failed: ' + IntToStr(countFail));
    Synmemo1.Lines.Add('prosent ok: ' + intToStr(pstOk) + '% ;' + 'prosent feil: ' + IntToStr(pstFail) + '%');
    Synmemo1.Lines.Add(StringOfChar('-',58));
  finally
    Comport1.Close();
  end;
end;

procedure TForm1.btnTestEepromClick(Sender: TObject);
var
  WriteBuffer :ByteArray;
  ReadBuffer  :ByteArray;
begin
  MessageDlg('Test av EEPROM kan opptil 5 minutter, Det VIL slette opptak i EEPROM' , mtWarning, [mbOK],0);

  //ABPMCommand(TestEEPROM,ReadBuffer);
  //EEPROM Test is extremely slow..... 4 mins....

  Comport1.Open();
  try
    FillChar(WriteBuffer,sizeOf(WriteBuffer),0);
    WriteBuffer[0] := 2;
    WriteBuffer[1] := Byte(TestEEPROM);
    WriteBuffer[23] := CalcChecksum(WriteBuffer);
    WriteBuffer[24] := 3;
    if Sync() then
      Comport1.Write(WriteBuffer,sizeOf(WriteBuffer));
    SynMemo1.Lines.Add('*VENNLIGST VENT*');
    sleep(5*60000);
    Comport1.Read(Readbuffer,sizeof(ReadBuffer));

    if ReadBuffer[2] = byte(58) then
    begin
      If ReadBuffer[3] = byte(1) then
        SynMemo1.Lines.Add('EEPROM test success!! ')
      else
        SynMemo1.Lines.Add('EEPROM test failed :( ');
    end;
  finally
    Comport1.Close();
  end;
end;

procedure TForm1.btnReadMeasurementsClick(Sender: TObject);
var
  WriteBuffer     :ByteArray;
  ReadBuffer      :ByteArray;
  measHi,measLo   :byte;
  Measurement     :TMeasurement;
  nytid           :TDateTime;
  tid             :string;
begin
  FillChar(WriteBuffer,SizeOf(WriteBuffer),0);

  measHi := HiByte(SpinEdit8.Value);
  measLo := LoByte(SpinEdit8.Value);

  WriteBuffer[0] := measHi;
  WriteBuffer[1] := measLo;

  ABPMCommand(ReadMeasurement,WriteBuffer,ReadBuffer);

  Measurement := TMeasurement.Create;
  try
    if ReadBuffer[2] = byte(ReadMeasurement) then
    begin
      Measurement.FSys := ReadBuffer[3] + 45;
      Measurement.FDia := ReadBuffer[4];
      Measurement.FMeanVal := ReadBuffer[5] + 35;
      Measurement.FPulse := ReadBuffer[6];
      Measurement.FHour := ReadBuffer[7];
      Measurement.FMin := ReadBuffer[8];
      Measurement.FDay := ReadBuffer[9];
      Measurement.FMonth := ReadBuffer[10];
      Measurement.FYear := ReadBuffer[11] + 2000;
      Measurement.FError := ReadBuffer[12];
      Measurement.FBatVolt := (ReadBuffer[13] * 2) * 10;

      LongTimeFormat := 'hh:mm';
      try
        nytid := EncodeDateTime(Measurement.Year,Measurement.Month,Measurement.Day,Measurement.Hour,Measurement.Min,0,0);
        tid := DateTimeToStr(nytid);
      Except
        On E: EConvertError do
        raise
      end;

      Synmemo1.Lines.Add('Sys: ' + IntToStr(Measurement.Sys));
      Synmemo1.Lines.Add('Dia: ' + IntToStr(Measurement.Dia));
      Synmemo1.Lines.Add('Mean val: ' + IntToStr(Measurement.Mean));
      Synmemo1.Lines.Add('Pulse: ' + IntToStr(Measurement.pulse));
      Synmemo1.Lines.Add('Tidspunkt: ' + tid);
      Synmemo1.Lines.Add('Error: ' + IntToStr(Measurement.Error));
      Synmemo1.Lines.Add('Batterinspenning: ' + IntToStr(Measurement.Volt));
    end
    else if ReadBuffer[2] = byte(NotAvailable) then
    begin
      Synmemo1.Lines.Add('Ingen måling funnet');
    end;
  finally
    Measurement.Free;
  end;
end;

procedure TForm1.btnReadHeaderClick(Sender: TObject);
var
  ReadBuffer    :ByteArray;
  WriteBuffer   :ByteArray;
  HeaderTxt     :string;
  blokk         :integer;
  I             :integer;
begin
  FillChar(WriteBuffer,sizeOf(WriteBuffer),0);
  blokk := 0;
  WriteBuffer[0] := blokk;
  ABPMCommand(ReadHeader,WriteBuffer,ReadBuffer);
  if ReadBuffer[2] = byte(HeaderRead) then
  begin
    for I := 6 to 21 do                                                         // Something odd for header data
      HeaderTxt := HeaderTxt + chr(ReadBuffer[i]);
 end;
 Synmemo1.Lines.Add('Board:' + HeaderTxt);
end;

procedure TForm1.btnBuzzerTestClick(Sender: TObject);
begin
  //ABPMCommand(156);
  ABPMCommand(TestBuzzer);
end;


Function TForm1.CalcChecksum(data: array of Byte):integer;
var
  i      :integer;
  check  :integer;
begin
  check := 0;
  for I := 1 to 22 do
    begin
      check := check + data[I]
    end;
  result := check;
end;


procedure TForm1.FormCreate(Sender: TObject);
const
  GUID_DEVICEINTERFACE_COMPORT: TGUID = '{86E0D1E0-8089-11D0-9CE4-08003E301F73}';
  Guid1: TGUID = '{9d7debbc-c85d-11d1-9eb4-006008c3a19a}';
begin
  DeviceNotifier := TDeviceNotifier.Create(Guid1);
  DeviceNotifier.FDeviceArrival := arrival;
  DeviceNotifier.FDeviceRemoval := removal;

  ShadowStrList := TStringlist.Create();
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  ShadowStrList.Free();
  DeviceNotifier.Free;
end;

procedure TForm1.btnCloseComportClick(Sender: TObject);
begin
  Comport1.Close;
end;

procedure TForm1.btnDelHeaderClick(Sender: TObject);
var
  btnValue  :integer;
begin
  btnValue := MessageDlg('Er du sikker på at du vil slette Header?', mtConfirmation, [mbYes, mbNo], 0);
  if btnValue = mrYes then
    ABPMCommand(DeleteHeader);
end;

procedure TForm1.btnDisplayTestClick(Sender: TObject);
begin
  ABPMCommand(TestDisplay);
  //ABPMCommand(150);
end;


procedure TForm1.lesSerieno(Sender: TObject);
{Manual claims serial should be 15 ASCII + CRC, seems false.
length of serial nr looks to be 6 bytes.}
var
  ReadBuffer  :ByteArray;
  SerialNo    :string;
  i           :integer;
begin
  //ABPMCommand(24,buff);
  ABPMCommand(ReadSerialNumber,ReadBuffer);
  if ReadBuffer[2] = byte(SerialNumber) then
  begin
    for I := 3 to 8 do
      SerialNo := SerialNo + chr(ReadBuffer[i]);
      SynMemo1.Lines.Add('Serienummeret er: ' + SerialNo);
  end;
end;


procedure TForm1.lesBatterispenning(Sender: TObject);
var
  ReadBuffer  :ByteArray;
  bVolt       :integer;
begin
  //ABPMCommand(51,buff);
  ABPMCommand(ReadBatteryVoltage,ReadBuffer);
  if ReadBuffer[2] = byte(BatteryVoltage) then
  begin
    bVolt :=  ReadBuffer[3]*256 + ReadBuffer[4];
    SynMemo1.Lines.Add('Batterispenningen er: ' + intToStr(bVolt) + ' mV');
  end
  else
    SynMemo1.Lines.Add('kunne ikke lese batterispenning i apparat!');
end;


procedure TForm1.btnButtonTestClick(Sender: TObject);
begin
  //ABPMCommand(154);
  ABPMCommand(TestButtons);
end;

procedure TForm1.lesVersjon(Sender: TObject);
var
  ReadBuffer :ByteArray;
  VerNo      :string;
  i          :integer;
begin
  //ABPMCommand(60,buff);
  ABPMCommand(ReadVersion,ReadBuffer);
  if ReadBuffer[2] = byte(Version) then
    for i := 3 to 23 do
    begin
      VerNo := VerNo + chr(ReadBuffer[I]);
    end;
  Synmemo1.Lines.Add('Versjon er: ' + VerNo);
end;

procedure TForm1.lesMinnespenning(Sender: TObject);
var
  ReadBuffer  :ByteArray;
  memVolt     :integer;
begin
  //ABPMCommand(55,buff);
  ABPMCommand(ReadMemoVoltage,ReadBuffer);
  //ABPMCommand(ReadMemVoltOld,ReadBuffer);
  if ReadBuffer[2] = byte(MemoVoltage) then
  begin
    memVolt := ReadBuffer[3]*256 + ReadBuffer[4];
    Synmemo1.Lines.Add('Minnespenningen er: ' + intToStr(memVolt) + ' mV');
  end
  else
    SynMemo1.Lines.Add('Kunne ikke lese spenningen på minnebatteri.');
end;

function TForm1.Sync():boolean;
var
  WriteBuffer  :byte;
  BytesRec     :integer;
  BytesSent    :integer;
  reply        :byte;
begin
  //sleep(50);
  //result := false;
  If Comport1.Connected then
  begin
    WriteBuffer := 5;                                 //Sync
    BytesRec := Comport1.Write(WriteBuffer,1);        //Sync er bare 1 byte
    sleep(20);                                         // Må sove litt .... :(  ...hmmm 
    BytesSent := Comport1.Read(reply,1);
  end;                                                //Check for ACK!!!
    result := (reply = 6);
end;

procedure TForm1.SynMemo1Click(Sender: TObject);
var
  LCoord    :TBufferCoord;
  LineText  :string;
  Startpos  :integer;
  Endpos    :integer;
begin
  Startpos := SynMemo1.SelStart;
  Endpos := SynMemo1.SelEnd;

  glLength := SynMemo1.SelLength;

  SynMemo1.GetPositionOfMouse(Lcoord);
  SynMemo1.CaretY := Lcoord.Line;
  SynMemo1.CaretX := 1;

  SynMemo1.SelStart := Startpos;
  //SynMemo1.SelLength := Length(SynMemo1.LineText);
  SynMemo1.SelEnd := Endpos;

  SynMemo1.SetFocus();
end;

procedure TForm1.SynMemo1SpecialLineColors(Sender: TObject; Line: Integer;
  var Special: Boolean; var FG, BG: TColor);
begin
  Special := true;
  if pos('[E010]',SynMemo1.lines[Line-1])>0 then
  begin
    BG := clWebDarkOrange;
    FG := clWebForestGreen;
  end;
  {if (line mod 4)=0 then
    BG := clWebLimeGreen;
  if (line mod 8) = 0 then
    BG := clWebSpringGreen;}  
end;

function TForm1.ABPMCommand(Command: TABPMCommand):boolean;
var
  WriteBuffer  :ByteArray;
  reply        :byte;
begin
  result := false;
  Comport1.Open();
  try
    FillChar(WriteBuffer,sizeOf(WriteBuffer),0);                                 //clearing buffer with 0's
    WriteBuffer[0] := 2;                                                         //Sx
    WriteBuffer[1] := byte(Command);                                             //command to send
    WriteBuffer[23] := CalcChecksum(WriteBuffer);                                //find checksum
    WriteBuffer[24] := 3;                                                        //ex
    if Sync() then                                                               //check sync
      Comport1.Write(WriteBuffer,25);                                            //write blokk to comport
    //sleep(20);                                                                   // some wait?! dunno why...
    Comport1.Read(reply,1);                                                      //read answer on comport

    result := (reply = 6);
    //This equals: 
    {if (reply = 6) then
      result := true; }
      
    if (reply = 21) then
      result := false;
  finally
    Comport1.Close();
  end;
end;

function Tform1.ABPMCommand(Command:TABPMCommand; out Data:ByteArray):boolean;
var
  WriteBuffer   :ByteArray;
  Readbuffer    :ByteArray;
begin
  result := false;
  try
  Comport1.Open();
    try
      FillChar(WriteBuffer,sizeOf(WriteBuffer),0);
      WriteBuffer[0] := 2;
      WriteBuffer[1] := Byte(Command);
      WriteBuffer[23] := CalcChecksum(WriteBuffer);
      WriteBuffer[24] := 3;
      if Sync() then
        Comport1.Write(WriteBuffer,sizeOf(WriteBuffer));
      Comport1.Read(Readbuffer,sizeof(ReadBuffer));
      Data := ReadBuffer;

    Except
      On E:EComPort do
        ShowMessage('Sjekk at apparat står i og riktig comport er valgt');
     end;

  finally
    Comport1.Close();
  end;
end;


function TForm1.ABPMCommand(Command: TABPMCommand; aData: ByteArray; out Data: ByteArray):boolean;
var
  WriteBuffer     :ByteArray;
  ReadBuffer      :ByteArray;            //array [0..24] of byte;
  i               :integer;
  len             :integer;
begin
  result := false;
  Comport1.Open();
  try
    FillChar(WriteBuffer,sizeOf(WriteBuffer),0);
    WriteBuffer[0] := 2;
    WriteBuffer[1] := byte(Command);

    len := length(aData);

    for i := 0 to len do
      WriteBuffer[i+2] := aData[i];

    WriteBuffer[23] := CalcChecksum(WriteBuffer);
    WriteBuffer[24] := 3;

    if Sync() then
      Comport1.Write(WriteBuffer,SizeOf(WriteBuffer));

    Comport1.Read(ReadBuffer,SizeOf(ReadBuffer));
    Data := ReadBuffer;
  finally
    Comport1.Close();
  end;
end;

function Tform1.ABPMCommand(Command:TABPMOldCommand; out Data:ByteArray):boolean;
var
  WriteBuffer   :ByteArray;
  Readbuffer    :ByteArray;
begin
  result := false;

  try
  Comport1.Open();
    try
      FillChar(WriteBuffer,sizeOf(WriteBuffer),0);
      WriteBuffer[0] := 2;
      WriteBuffer[1] := Byte(Command);
      WriteBuffer[23] := CalcChecksum(WriteBuffer);
      WriteBuffer[24] := 3;
      if Sync() then
        Comport1.Write(WriteBuffer,sizeOf(WriteBuffer));
      Comport1.Read(Readbuffer,sizeof(ReadBuffer));
      Data := ReadBuffer;

    Except
      On E:EComPort do
        ShowMessage('Sjekk at apparat står i og riktig comport er valgt');
     end;

  finally
    Comport1.Close();
  end;
end;

function TForm1.ABPMCommand(Command: TABPMOldCommand):boolean;
var
  WriteBuffer  :ByteArray;
  reply        :byte;
begin
  result := false;
  Comport1.Open();
  try
    FillChar(WriteBuffer,sizeOf(WriteBuffer),0);                                 //clearing buffer with 0's
    WriteBuffer[0] := 2;                                                         //Sx
    WriteBuffer[1] := byte(Command);                                             //command to send
    WriteBuffer[23] := CalcChecksum(WriteBuffer);                                //find checksum
    WriteBuffer[24] := 3;                                                        //ex
    if Sync() then                                                               //check sync
      Comport1.Write(WriteBuffer,25);                                            //write blokk to comport
    //sleep(20);                                                                   // some wait?! dunno why...
    Comport1.Read(reply,1);                                                      //read answer on comport
    if (reply = 6) then
      result := true;
    if (reply = 21) then
      result := false;
  finally
    Comport1.Close();
  end;
end;

function TForm1.ABPMCommand(Command: TABPMOldCommand; aData: ByteArray; out Data: ByteArray):boolean;
var
  WriteBuffer     :ByteArray;
  ReadBuffer      :ByteArray;            //array [0..24] of byte;
  i               :integer;
  len             :integer;
begin
  result := false;
  Comport1.Open();
  try
    FillChar(WriteBuffer,sizeOf(WriteBuffer),0);
    WriteBuffer[0] := 2;
    WriteBuffer[1] := byte(Command);

    len := length(aData);

    for i := 0 to len do
      WriteBuffer[i+2] := aData[i];

    WriteBuffer[23] := CalcChecksum(WriteBuffer);
    WriteBuffer[24] := 3;

    if Sync() then
      Comport1.Write(WriteBuffer,SizeOf(WriteBuffer));

    Comport1.Read(ReadBuffer,SizeOf(ReadBuffer));
    Data := ReadBuffer;
  finally
    Comport1.Close();
  end;
end;

end.
