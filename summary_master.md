# Summary Master
```
Private Sub Form_Sum_Master_Click()
```
- Inicia a definição de um procedimento (sub-rotina) que será executado quando o botão Form_Sum_Master for clicado.

_______

```
Dim RgR1, RgR2, RgR3 As Long
Dim WsOnhand, WsSumMaster, WsBOM As Worksheet

```
- Declara as variáveis RgR1, RgR2, e RgR3 como do tipo Long (números inteiros grandes) e as variáveis WsOnhand, WsSumMaster, e WsBOM como referências de planilhas (Worksheet).

_________

```
Application.Calculation = xlCalculationManual
```
- Define o modo de cálculo da aplicação Excel para manual, evitando recalcular automaticamente as fórmulas durante a execução da macro.

_______
```
Set WsSumMaster = Worksheets("Summary_Master")
Set WsBOM = Worksheets("BOM")
```
- Associa as variáveis WsSumMaster e WsBOM às planilhas "Summary_Master" e "BOM", respectivamente.
_______

```
WsSumMaster.Select
If ActiveSheet.FilterMode = True Then ActiveSheet.ShowAllData
RgR1 = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
Range("A4:S" & RgR1 + 50).EntireRow.Delete
```
- Seleciona a planilha WsSumMaster, remove todos os filtros ativos, identifica a última linha usada na coluna A, e apaga as linhas de A4 até RgR1+50.
______
```
WsBOM.Select
If ActiveSheet.FilterMode = True Then ActiveSheet.ShowAllData
RgR2 = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
ActiveSheet.Range("A1:AN" & RgR2).AutoFilter Field:=5, Criteria1:=Array("G", "M", "R"), Operator:=xlFilterValues
WsBOM.Range(Cells(2, "A"), Cells(RgR2, "A")).Copy WsSumMaster.Cells(4, "A")
```
- Seleciona a planilha WsBOM, remove todos os filtros ativos, encontra a última linha usada na coluna A, aplica um filtro na coluna E para mostrar apenas valores "G", "M" e "R", e copia a coluna A da WsBOM para WsSumMaster a partir da linha 4.
________
```
WsBOM.Range(Cells(2, "C"), Cells(RgR2, "D")).Copy WsSumMaster.Cells(4, "B")
WsBOM.Range(Cells(2, "N"), Cells(RgR2, "N")).Copy WsSumMaster.Cells(4, "D")
WsBOM.Range(Cells(2, "E"), Cells(RgR2, "E")).Copy WsSumMaster.Cells(4, "E")
WsBOM.Range(Cells(2, "K"), Cells(RgR2, "M")).Copy WsSumMaster.Cells(4, "F")
WsBOM.Range(Cells(2, "H"), Cells(RgR2, "H")).Copy WsSumMaster.Cells(4, "I")
```
- Copia várias colunas da WsBOM e cola nas respectivas colunas de WsSumMaster a partir da linha 4.
______
```
WsSumMaster.Select
If ActiveSheet.FilterMode = True Then ActiveSheet.ShowAllData
RgR1 = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
ActiveSheet.Range("A3:I" & RgR1).RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, 7, 8, 9), Header:=xlYes
```
- Seleciona a planilha WsSumMaster, remove todos os filtros ativos, encontra a última linha usada na coluna A, e remove duplicatas nas colunas de A até I.
_______
```
If ActiveSheet.FilterMode = True Then ActiveSheet.ShowAllData
RgR1 = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
ActiveSheet.Range("A4:I" & RgR1).Borders(xlDiagonalDown).LineStyle = xlNone
ActiveSheet.Range("A4:I" & RgR1).Borders(xlDiagonalUp).LineStyle = xlNone
ActiveSheet.Range("A4:I" & RgR1).Borders(xlEdgeLeft).LineStyle = xlNone
ActiveSheet.Range("A4:I" & RgR1).Borders(xlEdgeTop).LineStyle = xlNone
ActiveSheet.Range("A4:I" & RgR1).Borders(xlEdgeBottom).LineStyle = xlNone
ActiveSheet.Range("A4:I" & RgR1).Borders(xlEdgeRight).LineStyle = xlNone
ActiveSheet.Range("A4:I" & RgR1).Borders(xlInsideVertical).LineStyle = xlNone
ActiveSheet.Range("A4:I" & RgR1).Borders(xlInsideHorizontal).LineStyle = xlNone
[A4].Select
```
- Remove todos os filtros ativos, encontra a última linha usada na coluna A, remove todas as bordas das células de A4 até RgR1 em I, e seleciona a célula A4.
__________
```
WsSumMaster.Select
If ActiveSheet.FilterMode = True Then ActiveSheet.ShowAllData
RgR1 = WsSumMaster.Cells(Rows.Count, "A").End(xlUp).Row
Cells(RgR1, "S").Value = "Plan"
Cells(RgR1 + 1, "S").Value = "Delivery"
Cells(RgR1 + 2, "S").Value = "Balance"
For f = RgR1 To 5 Step -1
    Cells(f - 1, "S").Value = "Plan"
    Rows(f).Insert
    Cells(f, "S").Value = "Balance"
    Rows(f).Insert
    Cells(f, "S").Value = "Delivery"
Next f
```
- Seleciona a planilha WsSumMaster, remove todos os filtros ativos, encontra a última linha usada na coluna A, define valores na coluna S, e insere novas linhas com valores específicos em um loop.
_____
```
WsSumMaster.Select
If ActiveSheet.FilterMode = True Then ActiveSheet.ShowAllData
RgR1 = WsSumMaster.Cells(Rows.Count, "S").End(xlUp).Row
For f = RgR1 To 4 Step -1
    For g = 1 To 10 Step 1
        If Cells(f, "S") = "Balance" Then Cells(f, g) = Cells(f - 2, g)
        If Cells(f, "S") = "Delivery" Then Cells(f, g) = Cells(f - 1, g)
    Next g
    If Cells(f, "S") = "Plan" Then
        Range(Cells(f, "A"), Cells(f, "BH")).Interior.ThemeColor = xlThemeColorLight1
        Range(Cells(f, "A"), Cells(f, "BH")).Font.ThemeColor = xlThemeColorDark1
        Range(Cells(f, "J"), Cells(f, "BH")).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    ElseIf Cells(f, "S") = "Delivery" Then
        Range(Cells(f, "A"), Cells(f, "BH")).Interior.ThemeColor = xlThemeColorLight1
        Range(Cells(f, "S"), Cells(f, "BH")).Font.ThemeColor = xlThemeColorDark1
        Range(Cells(f, "J"), Cells(f, "BH")).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    ElseIf Cells(f, "S") = "Balance" Then
        Range(Cells(f, "A"), Cells(f, "R")).Interior.ThemeColor = xlThemeColorLight1
        Range(Cells(f, "S"), Cells(f, "BH")).Interior.Color = 255
        Range(Cells(f, "J"), Cells(f, "BH")).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
        Range(Cells(f, "S"), Cells(f, "BH")).Font.ThemeColor = xlThemeColorDark1
        Range(Cells(f, "S"), Cells(f, "BH")).Font.TintAndShade = 0
        Range(Cells(f, "A"), Cells(f, "BH")).Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range(Cells(f, "A"), Cells(f, "BH")).Borders(xlEdgeBottom).ThemeColor = 4
        Range(Cells(f, "A"), Cells(f, "BH")).Borders(xlEdgeBottom).TintAndShade = 0.599963377788629
    End If
Next f

```
_________
```
Range("J4:J" & RgR1).Borders(xlEdgeRight).LineStyle = xlContinuous
Range("J4:J" & RgR1).Borders(xlEdgeRight).ThemeColor = 4
Range("J4:J" & RgR1).Borders(xlEdgeRight).TintAndShade = 0.599963377788629
Range("J4:J" & RgR1).Borders(xlEdgeRight).Weight = xlThin
```
1. Range("J4:J" & RgR1): Seleciona o intervalo de células da coluna J, da linha 4 até a linha RgR1.

2. .Borders(xlEdgeRight).LineStyle = xlContinuous: Define a linha de borda direita dessas células como contínua.

3. .Borders(xlEdgeRight).ThemeColor = 4: Define a cor da borda direita usando um tema de cor específico.

4. .Borders(xlEdgeRight).TintAndShade = 0.599963377788629: Ajusta o tom e a matiz da cor da borda.

5. .Borders(xlEdgeRight).Weight = xlThin: Define a espessura da borda como fina.