unit Random;

interface

uses Math;

function RandomSentence: string;

implementation

uses Opts;

const  // may be customisable later
  MinWords = 3;  MaxWords = 8;
  CaseProb: array[0..1, 0..3] of Real = ((0.7, 0.1, 0.1, 0.1), (0.8, 0.1, 0.1, 0));
  CompoundProb = 0.05;
  PronounProb  = 0.05;

function RandomCase(Stress: Integer): string;
const vowels: array[0..2] of string = ('aeyioOuU', 'àèÌìòÒùÙ', 'áéıíóÓúÍÚ');
      suffs: array[0..1] of string = ('lRr', 'nm');
var p, r: Real;
    i, j: Integer;
begin
  Result:=vowels[Stress][System.Random(12) mod 8 +1];  
  for i:=0 to 1 do begin
    r:=System.Random;
    p:=0;
    for j:=0 to 3-i do begin
      p:=p+CaseProb[i][j];
      if r<p then begin
        if j>0 then Result:=Result+suffs[i][j];
        Break;
      end;
    end;
  end;
end; {RandomCase}

function RandomStem(Level: Integer): string;
const pronouns = 'wvzcj';
var r: Real;
begin
  r:=System.Random;
  if r<CompoundProb then Result:=RandomStem(Level)+RandomCase(0)+RandomStem(Level) else if r<CompoundProb+PronounProb then
  begin
    Result:=pronouns[System.Random(Level)+1];
    // type II pronouns
  end else
    with Options.Dictionary.Items do Result:='('+Names[System.Random(Count)]+')';
end; {RandomStem}

function RandomSentence: string;
const punctuation: array[0..2] of string = (' ', ', ', '.');
var level, n, f, p: Integer;
    s: array[0..1] of Integer;
begin
  Randomize;
  Result:='';
  level:=1;
  n:=0;
  repeat
    Inc(n);
    f:=Max(System.Random(level+1)+1, level-5);
    if (n<MinWords) and (f=1) then f:=2;
    s[0]:=0;  s[1]:=0;
    if f=1 then begin
      s[0]:=1;  p:=2;
    end else if level-f=-1 then begin
      s[0]:=System.Random(2)+1;  p:=0;
    end else begin
      s[1-((level-f) div 2 mod 2)]:=(level-f) mod 2 +1;
      p:=Ord(level-f>1);
    end;
    Result:=Result+RandomStem(level)+RandomCase(s[0]);
    if level>1 then Result:=Result+RandomCase(s[1]);
    Result:=Result+punctuation[p];
    level:=f;
  until (level=1) or (n>=MaxWords);
end; {RandomSentence}

end.
