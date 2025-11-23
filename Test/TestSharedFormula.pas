unit TestSharedFormula;

interface

uses
  DUnitX.TestFramework, Excel4Delphi.Stream;

type
  [TestFixture]
  SharedFormulaTest = class
  public
    // Sample Methods
    // Simple single Test
    [Test]
    [TestCase('TestA','0,0,1,1,1+A1,1+B2')]
    [TestCase('TestB','0,0,1,1,A1+1,B2+1')]
    [TestCase('TestC','0,0,1,1,A1+B2,B2+C3')]
    [TestCase('TestD','0,0,99,26,SUM(A1:C3),SUM(AA100:AC102)')]
    procedure Test(Top,Left,Row,Column: Integer; const Formula,Expected : String);
  end;

implementation

procedure SharedFormulaTest.Test(Top,Left,Row,Column: Integer; const Formula,Expected : String);
begin
  var SharedFormula := TZSharedFormula.Create(Top,Left,Formula);
  Assert.AreEqual(Expected,SharedFormula.Formula(Row,Column));
end;

initialization
  TDUnitX.RegisterTestFixture(SharedFormulaTest);

end.
