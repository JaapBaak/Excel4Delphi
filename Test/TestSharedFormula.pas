unit TestSharedFormula;

interface

uses
  DUnitX.TestFramework, Excel4Delphi, Excel4Delphi.Stream;

type
  [TestFixture]
  SharedFormulaTest = class
  public
    // Test whether the TZSharedFormula.Formula function correctly adjusts cell references
    [Test]
    [TestCase('TestA','0,0,1,1,1+A1,1+B2')]
    [TestCase('TestB','0,0,1,1,A1+1,B2+1')]
    [TestCase('TestC','0,0,1,1,A1+B2,B2+C3')]
    [TestCase('TestD','0,0,99,26,SUM(A1:C3),SUM(AA100:AC102)')]
    [TestCase('TestE','0,0,1,1,Sheet1!A1,Sheet1!B2')]
    [TestCase('TestF','0,0,99,26,SUM(Sheet1!A1:Sheet1!C3),SUM(Sheet1!AA100:Sheet1!AC102)')]
    procedure TestCellAdjustment(Top,Left,Row,Column: Integer; const Formula,Expected : String);
    // Test whether shared formula are correctly read from file
    [Test]
    procedure TestReadFromFile;
  end;

implementation

procedure SharedFormulaTest.TestCellAdjustment(Top,Left,Row,Column: Integer; const Formula,Expected : String);
begin
  var SharedFormula := TZSharedFormula.Create(Top,Left,Formula);
  Assert.AreEqual(Expected,SharedFormula.Formula(Row,Column));
end;

procedure SharedFormulaTest.TestReadFromFile;
begin
  var WorkBook := TZWorkBook.Create(nil);
  try
    // Read Excel-file containing shared formula, set in cell A3
    WorkBook.LoadFromFile('.\Data\SharedFormula.xlsx');
    // Check formula for cell A10
    Assert.AreEqual('A9+1',WorkBook.Sheets[0].Cell[0,9].Formula);
  finally
    WorkBook.Free;
  end;
end;

initialization
  TDUnitX.RegisterTestFixture(SharedFormulaTest);

end.
