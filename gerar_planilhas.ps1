# Script gerador da planilha OpenXML e CSV
Add-Type -AssemblyName System.IO.Compression.FileSystem

$baseDir = "c:\Users\yhago\Downloads\imposto"
$buildDir = "$baseDir\xlsx_build"
$xlsxPath = "$baseDir\analise_impostos_qualidade_vida_brasil.xlsx"
$csvPath  = "$baseDir\relatorio_impostos_qualidade_vida_brasil.csv"

if (Test-Path $buildDir) { Remove-Item -Recurse -Force $buildDir }
if (Test-Path $xlsxPath) { Remove-Item -Force $xlsxPath }

New-Item -ItemType Directory -Path "$buildDir\_rels" -Force | Out-Null
New-Item -ItemType Directory -Path "$buildDir\xl\_rels" -Force | Out-Null
New-Item -ItemType Directory -Path "$buildDir\xl\worksheets" -Force | Out-Null

# 1. [Content_Types].xml
$contentTypes = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet4.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet5.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>
'@
[System.IO.File]::WriteAllText("$buildDir\[Content_Types].xml", $contentTypes, [System.Text.Encoding]::UTF8)

# 2. _rels/.rels
$rootRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
'@
[System.IO.File]::WriteAllText("$buildDir\_rels\.rels", $rootRels, [System.Text.Encoding]::UTF8)

# 3. xl/_rels/workbook.xml.rels
$wbRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet4.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet5.xml"/>
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
'@
[System.IO.File]::WriteAllText("$buildDir\xl\_rels\workbook.xml.rels", $wbRels, [System.Text.Encoding]::UTF8)

# 4. xl/workbook.xml
$workbook = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Resumo_Executivo" sheetId="1" r:id="rId1"/>
    <sheet name="Carga_Tributaria_Historica" sheetId="2" r:id="rId2"/>
    <sheet name="Educacao_IDEB_PISA" sheetId="3" r:id="rId3"/>
    <sheet name="Poder_de_Compra_DIEESE" sheetId="4" r:id="rId4"/>
    <sheet name="Destino_Impostos_IRBES" sheetId="5" r:id="rId5"/>
  </sheets>
</workbook>
'@
[System.IO.File]::WriteAllText("$buildDir\xl\workbook.xml", $workbook, [System.Text.Encoding]::UTF8)

# 5. xl/styles.xml
$styles = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="10"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="12"/><color rgb="FF0D233A"/><name val="Calibri"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F497D"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFDCE6F1"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFD3D3D3"/></left>
      <right style="thin"><color rgb="FFD3D3D3"/></right>
      <top style="thin"><color rgb="FFD3D3D3"/></top>
      <bottom style="thin"><color rgb="FFD3D3D3"/></bottom>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="4">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
  </cellXfs>
</styleSheet>
'@
[System.IO.File]::WriteAllText("$buildDir\xl\styles.xml", $styles, [System.Text.Encoding]::UTF8)

# Helper function to generate cell XML
function Get-CellXml([string]$col, [int]$row, [string]$val, [int]$style = 2, [bool]$isNum = $false) {
    $ref = "$col$row"
    if ($isNum) {
        return "<c r=`"$ref`" s=`"$style`"><v>$val</v></c>"
    } else {
        $clean = [System.Security.SecurityElement]::Escape($val)
        return "<c r=`"$ref`" s=`"$style`" t=`"inlineStr`"><is><t>$clean</t></is></c>"
    }
}

# ----------------- SHEET 1: Resumo_Executivo -----------------
$s1 = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>
    <col min="1" max="1" width="36"/>
    <col min="2" max="2" width="22"/>
    <col min="3" max="3" width="22"/>
    <col min="4" max="4" width="22"/>
    <col min="5" max="5" width="22"/>
    <col min="6" max="6" width="30"/>
  </cols>
  <sheetData>
    <row r="1">
      <c r="A1" s="3" t="inlineStr"><is><t>COMPARATIVO HISTORICO: EVOLUCAO DA CARGA TRIBUTARIA VS RETORNO EM QUALIDADE DE VIDA</t></is></c>
    </row>
    <row r="2">
      <c r="A2" s="1" t="inlineStr"><is><t>Indicador Socioeconomico</t></is></c>
      <c r="B2" s="1" t="inlineStr"><is><t>Era 1995-1999 (Menos Imposto)</t></is></c>
      <c r="C2" s="1" t="inlineStr"><is><t>Era 2000-2010 (Expansao)</t></is></c>
      <c r="D2" s="1" t="inlineStr"><is><t>Era 2011-2019 (Crise e Platô)</t></is></c>
      <c r="E2" s="1" t="inlineStr"><is><t>Era 2020-2025 (Mais Imposto / Atual)</t></is></c>
      <c r="F2" s="1" t="inlineStr"><is><t>Avaliacao do Retorno / Conclusao</t></is></c>
    </row>
    <row r="3">
      <c r="A3" s="2" t="inlineStr"><is><t>Carga Tributaria Bruta Media (% PIB)</t></is></c>
      <c r="B3" s="2" t="inlineStr"><is><t>26,7% a 28,0%</t></is></c>
      <c r="C3" s="2" t="inlineStr"><is><t>30,0% a 32,8%</t></is></c>
      <c r="D3" s="2" t="inlineStr"><is><t>32,1% a 32,5%</t></is></c>
      <c r="E3" s="2" t="inlineStr"><is><t>32,4% a 33,9%</t></is></c>
      <c r="F3" s="2" t="inlineStr"><is><t>Aumento estrutural de cerca de 7 p.p. do PIB</t></is></c>
    </row>
    <row r="4">
      <c r="A4" s="2" t="inlineStr"><is><t>Dias Trabalhados no Ano p/ Pagar Tributos (IBPT)</t></is></c>
      <c r="B4" s="2" t="inlineStr"><is><t>102 dias</t></is></c>
      <c r="C4" s="2" t="inlineStr"><is><t>138 dias</t></is></c>
      <c r="D4" s="2" t="inlineStr"><is><t>151 dias</t></is></c>
      <c r="E4" s="2" t="inlineStr"><is><t>149 dias (~5 meses do ano)</t></is></c>
      <c r="F4" s="2" t="inlineStr"><is><t>Cidadao trabalha 47 dias a mais por ano so p/ tributos</t></is></c>
    </row>
    <row r="5">
      <c r="A5" s="2" t="inlineStr"><is><t>Gasto Publico em Educacao (% do PIB)</t></is></c>
      <c r="B5" s="2" t="inlineStr"><is><t>3,5% do PIB</t></is></c>
      <c r="C5" s="2" t="inlineStr"><is><t>3,9% a 5,0% do PIB</t></is></c>
      <c r="D5" s="2" t="inlineStr"><is><t>5,5% a 5,8% do PIB</t></is></c>
      <c r="E5" s="2" t="inlineStr"><is><t>5,6% a 6,0% do PIB</t></is></c>
      <c r="F5" s="2" t="inlineStr"><is><t>Gasto % PIB superou a media da OCDE (~4,9%)</t></is></c>
    </row>
    <row r="6">
      <c r="A6" s="2" t="inlineStr"><is><t>PISA - Matematica (Proficiencia OCDE)</t></is></c>
      <c r="B6" s="2" t="inlineStr"><is><t>Nao medido antes de 2003</t></is></c>
      <c r="C6" s="2" t="inlineStr"><is><t>356 (2003) -> 386 (2009)</t></is></c>
      <c r="D6" s="2" t="inlineStr"><is><t>389 (2012) -> 384 (2018)</t></is></c>
      <c r="E6" s="2" t="inlineStr"><is><t>379 (2022)</t></is></c>
      <c r="F6" s="2" t="inlineStr"><is><t>Estagnado ha 15 anos; ~70% abaixo do nivel basico</t></is></c>
    </row>
    <row r="7">
      <c r="A7" s="2" t="inlineStr"><is><t>PISA - Leitura (Proficiencia OCDE)</t></is></c>
      <c r="B7" s="2" t="inlineStr"><is><t>396 (2000)</t></is></c>
      <c r="C7" s="2" t="inlineStr"><is><t>403 (2003) -> 412 (2009)</t></is></c>
      <c r="D7" s="2" t="inlineStr"><is><t>407 (2015) -> 413 (2018)</t></is></c>
      <c r="E7" s="2" t="inlineStr"><is><t>410 (2022)</t></is></c>
      <c r="F7" s="2" t="inlineStr"><is><t>Media OCDE e ~476; Brasil estacionado na lanterna</t></is></c>
    </row>
    <row r="8">
      <c r="A8" s="2" t="inlineStr"><is><t>IDEB - Ensino Medio (Escala 0 a 10)</t></is></c>
      <c r="B8" s="2" t="inlineStr"><is><t>Sem IDEB antes de 2005</t></is></c>
      <c r="C8" s="2" t="inlineStr"><is><t>3,4 (2005) -> 3,6 (2009)</t></is></c>
      <c r="D8" s="2" t="inlineStr"><is><t>3,7 (2013) -> 4,2 (2019)</t></is></c>
      <c r="E8" s="2" t="inlineStr"><is><t>4,3 (2023)</t></is></c>
      <c r="F8" s="2" t="inlineStr"><is><t>Ficou muito longe da meta de 5,2 estipulada p/ 2022</t></is></c>
    </row>
    <row r="9">
      <c r="A9" s="2" t="inlineStr"><is><t>Horas Trabalhadas p/ Comprar Cesta Basica (DIEESE)</t></is></c>
      <c r="B9" s="2" t="inlineStr"><is><t>143 horas (1995)</t></is></c>
      <c r="C9" s="2" t="inlineStr"><is><t>122h (2005) -> 108h (2010)</t></is></c>
      <c r="D9" s="2" t="inlineStr"><is><t>95h (2014) -> 102h (2018)</t></is></c>
      <c r="E9" s="2" t="inlineStr"><is><t>122h (2021) -> 107h (2025)</t></is></c>
      <c r="F9" s="2" t="inlineStr"><is><t>Consome mais de 50% do salario minimo liquido</t></is></c>
    </row>
    <row r="10">
      <c r="A10" s="2" t="inlineStr"><is><t>Indice de Retorno ao Bem-Estar - IRBES (IBPT)</t></is></c>
      <c r="B10" s="2" t="inlineStr"><is><t>Posicao Baixa</t></is></c>
      <c r="C10" s="2" t="inlineStr"><is><t>30º de 30 paises (Lanterna)</t></is></c>
      <c r="D10" s="2" t="inlineStr"><is><t>30º de 30 paises (Lanterna)</t></is></c>
      <c r="E10" s="2" t="inlineStr"><is><t>30º de 30 paises (Ultimo Lugar)</t></is></c>
      <c r="F10" s="2" t="inlineStr"><is><t>Pior retorno de servicos publicos entre os 30 maiores</t></is></c>
    </row>
    <row r="11">
      <c r="A11" s="2" t="inlineStr"><is><t>Investimento Publico Federal em Infraestrutura (% PIB)</t></is></c>
      <c r="B11" s="2" t="inlineStr"><is><t>1,8% do PIB</t></is></c>
      <c r="C11" s="2" t="inlineStr"><is><t>1,2% a 1,4% do PIB</t></is></c>
      <c r="D11" s="2" t="inlineStr"><is><t>0,7% a 0,9% do PIB</t></is></c>
      <c r="E11" s="2" t="inlineStr"><is><t>0,3% a 0,6% do PIB</t></is></c>
      <c r="F11" s="2" t="inlineStr"><is><t>Colapso do investimento: dinheiro vai para custeio e divida</t></is></c>
    </row>
    <row r="12">
      <c r="A12" s="2" t="inlineStr"><is><t>Indice de Desenvolvimento Humano - IDH (PNUD)</t></is></c>
      <c r="B12" s="2" t="inlineStr"><is><t>0,653 (1995)</t></is></c>
      <c r="C12" s="2" t="inlineStr"><is><t>0,704 (2005) -> 0,726 (2010)</t></is></c>
      <c r="D12" s="2" t="inlineStr"><is><t>0,757 (2015) -> 0,766 (2019)</t></is></c>
      <c r="E12" s="2" t="inlineStr"><is><t>0,760 (2022) -> 0,765 (2023)</t></is></c>
      <c r="F12" s="2" t="inlineStr"><is><t>Avanco puxado por longevidade; educacao e renda lentos</t></is></c>
    </row>
  </sheetData>
</worksheet>
'@
[System.IO.File]::WriteAllText("$buildDir\xl\worksheets\sheet1.xml", $s1, [System.Text.Encoding]::UTF8)

# ----------------- SHEET 2: Carga_Tributaria_Historica -----------------
$s2Lines = @(
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
  '  <cols>',
  '    <col min="1" max="1" width="10"/>',
  '    <col min="2" max="2" width="22"/>',
  '    <col min="3" max="3" width="26"/>',
  '    <col min="4" max="4" width="24"/>',
  '    <col min="5" max="5" width="28"/>',
  '  </cols>',
  '  <sheetData>',
  '    <row r="1">',
  '      <c r="A1" s="1" t="inlineStr"><is><t>Ano</t></is></c>',
  '      <c r="B1" s="1" t="inlineStr"><is><t>Carga Tributaria (% PIB)</t></is></c>',
  '      <c r="C1" s="1" t="inlineStr"><is><t>Arrecadacao Total (R$ Bilhoes)</t></is></c>',
  '      <c r="D1" s="1" t="inlineStr"><is><t>Dias Trabalhados p/ Imposto</t></is></c>',
  '      <c r="E1" s="1" t="inlineStr"><is><t>Contexto Fiscal e Economico</t></is></c>',
  '    </row>'
)

$historicoCarga = @(
  @{ Ano=1995; Carga="26.7"; Arrec="173.2"; Dias=102; Ctx="Implantacao do Plano Real e fim da hiperinflacao" },
  @{ Ano=1996; Carga="25.8"; Arrec="201.5"; Dias=101; Ctx="Estabilizacao inicial da moeda" },
  @{ Ano=1997; Carga="26.8"; Arrec="238.1"; Dias=103; Ctx="Crise Asiatica; ajuste fiscal inicial" },
  @{ Ano=1998; Carga="27.8"; Arrec="262.3"; Dias=105; Ctx="Crise da Russia; elevacao da CPMF e Cofins" },
  @{ Ano=1999; Carga="29.2"; Arrec="312.4"; Dias=112; Ctx="Desvalorizacao cambial e novo regime fiscal" },
  @{ Ano=2000; Carga="30.0"; Arrec="357.2"; Dias=128; Ctx="Lei de Responsabilidade Fiscal e alta de aliquotas" },
  @{ Ano=2001; Carga="31.1"; Arrec="406.8"; Dias=130; Ctx="Crise energetica (apagao) e manutencao da CPMF" },
  @{ Ano=2002; Carga="32.1"; Arrec="476.5"; Dias=133; Ctx="Criacao do PIS nao-cumulativo" },
  @{ Ano=2003; Carga="31.4"; Arrec="535.1"; Dias=134; Ctx="Transicao de governo e controle inflacionario" },
  @{ Ano=2004; Carga="32.0"; Arrec="622.8"; Dias=136; Ctx="Criacao da Cofins nao-cumulativa (alta de tributos)" },
  @{ Ano=2005; Carga="32.8"; Arrec="715.4"; Dias=138; Ctx="Boom internacional de commodities" },
  @{ Ano=2006; Carga="32.8"; Arrec="794.1"; Dias=140; Ctx="Crescimento economico e formalizacao do emprego" },
  @{ Ano=2007; Carga="33.2"; Arrec="901.3"; Dias=142; Ctx="Fim da CPMF no final do ano; PAC 1" },
  @{ Ano=2008; Carga="33.6"; Arrec="1038.2"; Dias=145; Ctx="Pico de arrecadacao antes da crise global do subprime" },
  @{ Ano=2009; Carga="32.1"; Arrec="1092.5"; Dias=143; Ctx="Desonerações de IPI para combater a crise mundial" },
  @{ Ano=2010; Carga="32.4"; Arrec="1288.7"; Dias=148; Ctx="Forte crescimento do PIB (+7,5%) e arrecadacao recorde" },
  @{ Ano=2011; Carga="32.9"; Arrec="1476.3"; Dias=149; Ctx="Nova Matriz Economica; inicio de subsidios setoriais" },
  @{ Ano=2012; Carga="32.7"; Arrec="1612.4"; Dias=149; Ctx="Desoneracao ampla da folha de pagamento" },
  @{ Ano=2013; Carga="32.5"; Arrec="1762.1"; Dias=150; Ctx="Jornadas de Junho: sociedade protesta contra impostos e servicos ruins" },
  @{ Ano=2014; Carga="32.1"; Arrec="1855.8"; Dias=150; Ctx="Deterioracao fiscal e inicio de deficit primario" },
  @{ Ano=2015; Carga="32.1"; Arrec="1932.6"; Dias=151; Ctx="Grave recessao economica e ajuste fiscal com reonerações" },
  @{ Ano=2016; Carga="32.3"; Arrec="2005.1"; Dias=152; Ctx="Recessao persistente e aprovacao do Teto de Gastos (EC 95)" },
  @{ Ano=2017; Carga="32.4"; Arrec="2158.4"; Dias=153; Ctx="Recuperacao lenta e Reforma Trabalhista" },
  @{ Ano=2018; Carga="32.5"; Arrec="2320.6"; Dias=153; Ctx="Greve dos caminhoneiros; subsidio ao diesel" },
  @{ Ano=2019; Carga="32.5"; Arrec="2475.9"; Dias=153; Ctx="Aprovacao da Reforma da Previdencia (EC 103)" },
  @{ Ano=2020; Carga="31.8"; Arrec="2450.2"; Dias=151; Ctx="Pandemia de COVID-19; auxilio emergencial e choque economico" },
  @{ Ano=2021; Carga="33.9"; Arrec="3041.5"; Dias=149; Ctx="Reabertura pos-pandemia e alta expressiva de commodities/inflacao" },
  @{ Ano=2022; Carga="33.7"; Arrec="3358.2"; Dias=149; Ctx="Selic a 13,75% elevando arrecadacao sobre capital financeiro" },
  @{ Ano=2023; Carga="32.4"; Arrec="3421.8"; Dias=147; Ctx="Novo Arcabouco Fiscal e aprovacao da Reforma Tributaria do Consumo" },
  @{ Ano=2024; Carga="32.3"; Arrec="3680.0"; Dias=149; Ctx="Reoneracao de combustiveis e medidas de fechamento de brechas fiscais" },
  @{ Ano=2025; Carga="32.4"; Arrec="3950.0"; Dias=149; Ctx="Arrecadacao nominal recorde proxima a R$ 4 trilhoes (estimativa)" }
)

$rIdx = 2
foreach ($item in $historicoCarga) {
    $rIdx++
    $s2Lines += "    <row r=`"$rIdx`">"
    $s2Lines += "      <c r=`"A$rIdx`" s=`"2`"><v>$($item.Ano)</v></c>"
    $s2Lines += "      <c r=`"B$rIdx`" s=`"2`"><v>$($item.Carga)</v></c>"
    $s2Lines += "      <c r=`"C$rIdx`" s=`"2`"><v>$($item.Arrec)</v></c>"
    $s2Lines += "      <c r=`"D$rIdx`" s=`"2`"><v>$($item.Dias)</v></c>"
    $cleanCtx = [System.Security.SecurityElement]::Escape($item.Ctx)
    $s2Lines += "      <c r=`"E$rIdx`" s=`"2`" t=`"inlineStr`"><is><t>$cleanCtx</t></is></c>"
    $s2Lines += "    </row>"
}
$s2Lines += '  </sheetData>'
$s2Lines += '</worksheet>'
[System.IO.File]::WriteAllLines("$buildDir\xl\worksheets\sheet2.xml", $s2Lines, [System.Text.Encoding]::UTF8)

# ----------------- SHEET 3: Educacao_IDEB_PISA -----------------
$s3 = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>
    <col min="1" max="1" width="12"/>
    <col min="2" max="2" width="22"/>
    <col min="3" max="3" width="20"/>
    <col min="4" max="4" width="20"/>
    <col min="5" max="5" width="20"/>
    <col min="6" max="6" width="20"/>
    <col min="7" max="7" width="20"/>
    <col min="8" max="8" width="28"/>
  </cols>
  <sheetData>
    <row r="1">
      <c r="A1" s="3" t="inlineStr"><is><t>PAINEL DE DADOS DA EDUCACAO: IDEB (INEP) E PISA (OCDE) VS GASTO PUBLICO</t></is></c>
    </row>
    <row r="2">
      <c r="A2" s="1" t="inlineStr"><is><t>Ano</t></is></c>
      <c r="B2" s="1" t="inlineStr"><is><t>Gasto Educacao (% PIB)</t></is></c>
      <c r="C2" s="1" t="inlineStr"><is><t>IDEB Anos Iniciais</t></is></c>
      <c r="D2" s="1" t="inlineStr"><is><t>IDEB Anos Finais</t></is></c>
      <c r="E2" s="1" t="inlineStr"><is><t>IDEB Ensino Medio</t></is></c>
      <c r="F2" s="1" t="inlineStr"><is><t>PISA Matematica</t></is></c>
      <c r="G2" s="1" t="inlineStr"><is><t>PISA Leitura</t></is></c>
      <c r="H2" s="1" t="inlineStr"><is><t>PISA Ciencias / Observacao</t></is></c>
    </row>
    <row r="3">
      <c r="A3" s="2"><v>2000</v></c>
      <c r="B3" s="2"><v>3.5</v></c>
      <c r="C3" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="D3" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="E3" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="F3" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="G3" s="2"><v>396</v></c>
      <c r="H3" s="2" t="inlineStr"><is><t>Primeiro ciclo do PISA (Brasil 32º de 32)</t></is></c>
    </row>
    <row r="4">
      <c r="A4" s="2"><v>2003</v></c>
      <c r="B4" s="2"><v>3.7</v></c>
      <c r="C4" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="D4" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="E4" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="F4" s="2"><v>356</v></c>
      <c r="G4" s="2"><v>403</v></c>
      <c r="H3" s="2" t="inlineStr"><is><t>Matematica inserida oficialmente</t></is></c>
    </row>
    <row r="5">
      <c r="A5" s="2"><v>2005</v></c>
      <c r="B5" s="2"><v>3.9</v></c>
      <c r="C5" s="2"><v>3.8</v></c>
      <c r="D5" s="2"><v>3.5</v></c>
      <c r="E5" s="2"><v>3.4</v></c>
      <c r="F5" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="G5" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="H5" s="2" t="inlineStr"><is><t>Criacao do indicador oficial IDEB</t></is></c>
    </row>
    <row r="6">
      <c r="A6" s="2"><v>2006</v></c>
      <c r="B6" s="2"><v>4.1</v></c>
      <c r="C6" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="D6" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="E6" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="F6" s="2"><v>370</v></c>
      <c r="G6" s="2"><v>393</v></c>
      <c r="H6" s="2"><v>390</v></c>
    </row>
    <row r="7">
      <c r="A7" s="2"><v>2007</v></c>
      <c r="B7" s="2"><v>4.5</v></c>
      <c r="C7" s="2"><v>4.2</v></c>
      <c r="D7" s="2"><v>3.8</v></c>
      <c r="E7" s="2"><v>3.5</v></c>
      <c r="F7" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="G7" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="H7" s="2" t="inlineStr"><is><t>Criacao do FUNDEB</t></is></c>
    </row>
    <row r="8">
      <c r="A8" s="2"><v>2009</v></c>
      <c r="B8" s="2"><v>4.8</v></c>
      <c r="C8" s="2"><v>4.6</v></c>
      <c r="D8" s="2"><v>4.0</v></c>
      <c r="E8" s="2"><v>3.6</v></c>
      <c r="F8" s="2"><v>386</v></c>
      <c r="G8" s="2"><v>412</v></c>
      <c r="H8" s="2"><v>405</v></c>
    </row>
    <row r="9">
      <c r="A9" s="2"><v>2011</v></c>
      <c r="B9" s="2"><v>5.1</v></c>
      <c r="C9" s="2"><v>5.0</v></c>
      <c r="D9" s="2"><v>4.1</v></c>
      <c r="E9" s="2"><v>3.7</v></c>
      <c r="F9" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="G9" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="H9" s="2" t="inlineStr"><is><t>Avanco nos Anos Iniciais do fundamental</t></is></c>
    </row>
    <row r="10">
      <c r="A10" s="2"><v>2012</v></c>
      <c r="B10" s="2"><v>5.3</v></c>
      <c r="C10" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="D10" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="E10" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="F10" s="2"><v>389</v></c>
      <c r="G10" s="2"><v>407</v></c>
      <c r="H10" s="2"><v>402</v></c>
    </row>
    <row r="11">
      <c r="A11" s="2"><v>2013</v></c>
      <c r="B11" s="2"><v>5.4</v></c>
      <c r="C11" s="2"><v>5.2</v></c>
      <c r="D11" s="2"><v>4.2</v></c>
      <c r="E11" s="2"><v>3.7</v></c>
      <c r="F11" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="G11" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="H11" s="2" t="inlineStr"><is><t>Estagnacao no Ensino Medio</t></is></c>
    </row>
    <row r="12">
      <c r="A12" s="2"><v>2015</v></c>
      <c r="B12" s="2"><v>5.5</v></c>
      <c r="C12" s="2"><v>5.5</v></c>
      <c r="D12" s="2"><v>4.5</v></c>
      <c r="E12" s="2"><v>3.7</v></c>
      <c r="F12" s="2"><v>377</v></c>
      <c r="G12" s="2"><v>407</v></c>
      <c r="H12" s="2"><v>401</v></c>
    </row>
    <row r="13">
      <c r="A13" s="2"><v>2017</v></c>
      <c r="B13" s="2"><v>5.7</v></c>
      <c r="C13" s="2"><v>5.8</v></c>
      <c r="D13" s="2"><v>4.7</v></c>
      <c r="E13" s="2"><v>3.8</v></c>
      <c r="F13" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="G13" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="H13" s="2" t="inlineStr"><is><t>Aprovacao da Reforma do Ensino Medio</t></is></c>
    </row>
    <row r="14">
      <c r="A14" s="2"><v>2018</v></c>
      <c r="B14" s="2"><v>5.6</v></c>
      <c r="C14" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="D14" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="E14" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="F14" s="2"><v>384</v></c>
      <c r="G14" s="2"><v>413</v></c>
      <c r="H14" s="2"><v>404</v></c>
    </row>
    <row r="15">
      <c r="A15" s="2"><v>2019</v></c>
      <c r="B15" s="2"><v>5.7</v></c>
      <c r="C15" s="2"><v>5.9</v></c>
      <c r="D15" s="2"><v>4.9</v></c>
      <c r="E15" s="2"><v>4.2</v></c>
      <c r="F15" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="G15" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="H15" s="2" t="inlineStr"><is><t>Leve alta no EM com distorcao de fluxo</t></is></c>
    </row>
    <row r="16">
      <c r="A16" s="2"><v>2021</v></c>
      <c r="B16" s="2"><v>5.8</v></c>
      <c r="C16" s="2"><v>5.8</v></c>
      <c r="D16" s="2"><v>4.9</v></c>
      <c r="E16" s="2"><v>4.2</v></c>
      <c r="F16" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="G16" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="H16" s="2" t="inlineStr"><is><t>Impactos do fechamento de escolas na pandemia</t></is></c>
    </row>
    <row r="17">
      <c r="A17" s="2"><v>2022</v></c>
      <c r="B17" s="2"><v>5.6</v></c>
      <c r="C17" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="D17" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="E17" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="F17" s="2"><v>379</v></c>
      <c r="G17" s="2"><v>410</v></c>
      <c r="H17" s="2"><v>403</v></c>
    </row>
    <row r="18">
      <c r="A18" s="2"><v>2023</v></c>
      <c r="B18" s="2"><v>5.7</v></c>
      <c r="C18" s="2"><v>6.0</v></c>
      <c r="D18" s="2"><v>5.0</v></c>
      <c r="E18" s="2"><v>4.3</v></c>
      <c r="F18" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="G18" s="2" t="inlineStr"><is><t>-</t></is></c>
      <c r="H18" s="2" t="inlineStr"><is><t>Anos iniciais bateu meta; EM longe da meta (5,2)</t></is></c>
    </row>
  </sheetData>
</worksheet>
'@
[System.IO.File]::WriteAllText("$buildDir\xl\worksheets\sheet3.xml", $s3, [System.Text.Encoding]::UTF8)

# ----------------- SHEET 4: Poder_de_Compra_DIEESE -----------------
$s4Lines = @(
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
  '  <cols>',
  '    <col min="1" max="1" width="10"/>',
  '    <col min="2" max="2" width="22"/>',
  '    <col min="3" max="3" width="24"/>',
  '    <col min="4" max="4" width="26"/>',
  '    <col min="5" max="5" width="26"/>',
  '    <col min="6" max="6" width="30"/>',
  '  </cols>',
  '  <sheetData>',
  '    <row r="1">',
  '      <c r="A1" s="3" t="inlineStr"><is><t>PODER DE COMPRA: SALARIO MINIMO VS CESTA BASICA (DIEESE) E TRIBUTACAO NO CONSUMO</t></is></c>',
  '    </row>',
  '    <row r="2">',
  '      <c r="A2" s="1" t="inlineStr"><is><t>Ano</t></is></c>',
  '      <c r="B2" s="1" t="inlineStr"><is><t>Salario Minimo Nominal (R$)</t></is></c>',
  '      <c r="C2" s="1" t="inlineStr"><is><t>Custo Cesta Basica DIEESE (R$)</t></is></c>',
  '      <c r="D2" s="1" t="inlineStr"><is><t>Horas Trabalhadas p/ Cesta</t></is></c>',
  '      <c r="E2" s="1" t="inlineStr"><is><t>% do Minimo Consumido</t></is></c>',
  '      <c r="F2" s="1" t="inlineStr"><is><t>Impacto da Tributacao Indireta</t></is></c>',
  '    </row>'
)

$dadosPoderCompra = @(
  @{ Ano=1995; SM=100;  CB=65.1;  Horas=143; Perc="65.1%"; Imp="Tributos no consumo representavam ~45% da arrecadacao" },
  @{ Ano=1998; SM=130;  CB=92.3;  Horas=156; Perc="71.0%"; Imp="Impacto da inflacao residual e impostos em cascata" },
  @{ Ano=2000; SM=151;  CB=115.4; Horas=135; Perc="61.4%"; Imp="Inicio de politica de recuperacao real do salario minimo" },
  @{ Ano=2002; SM=200;  CB=140.2; Horas=123; Perc="55.9%"; Imp="Alta do ICMS e PIS sobre cesta de servicos" },
  @{ Ano=2004; SM=260;  CB=162.8; Horas=120; Perc="54.5%"; Imp="Nova lei da Cofins encarece intermediarios" },
  @{ Ano=2006; SM=350;  CB=166.5; Horas=105; Perc="47.6%"; Imp="Ganho real expressivo do salario minimo" },
  @{ Ano=2008; SM=415;  CB=224.2; Horas=119; Perc="54.0%"; Imp="Choque global de alimentos (commodities)" },
  @{ Ano=2010; SM=510;  CB=251.6; Horas=108; Perc="49.3%"; Imp="Poder de compra em expansao no pico de crescimento" },
  @{ Ano=2012; SM=622;  CB=276.5; Horas=98;  Perc="44.5%"; Imp="Melhor patamar da serie (menos de 100 horas de trabalho)" },
  @{ Ano=2014; SM=724;  CB=313.2; Horas=95;  Perc="43.3%"; Imp="Menor comprometimento do salario com a cesta basica" },
  @{ Ano=2016; SM=880;  CB=438.1; Horas=110; Perc="49.8%"; Imp="Inflacao atinge 10,67% (IPCA) corroendo poder de compra" },
  @{ Ano=2018; SM=954;  CB=442.5; Horas=102; Perc="46.4%"; Imp="Desinflacao e estagnacao de rendimentos" },
  @{ Ano=2020; SM=1045; CB=580.4; Horas=112; Perc="55.5%"; Imp="Pandemia e disparada nos precos de arroz, oleo e carne" },
  @{ Ano=2021; SM=1100; CB=692.3; Horas=122; Perc="62.9%"; Imp="Forte inflacao de alimentos; perda severa de poder de compra" },
  @{ Ano=2022; SM=1212; CB=791.3; Horas=119; Perc="65.3%"; Imp="Cesta atinge patamar recorde de R$ 791 em Sao Paulo" },
  @{ Ano=2023; SM=1320; CB=772.5; Horas=109; Perc="58.5%"; Imp="Deflacao em alimentos no 2º semestre alivia pressao" },
  @{ Ano=2024; SM=1412; CB=825.0; Horas=108; Perc="58.4%"; Imp="Aumento real do salario minimo e inflacao moderada" },
  @{ Ano=2025; SM=1518; CB=845.0; Horas=107; Perc="55.7%"; Imp="Mais de 55% do piso salarial ainda gasto apenas com comida" }
)

$r4 = 2
foreach ($d in $dadosPoderCompra) {
    $r4++
    $s4Lines += "    <row r=`"$r4`">"
    $s4Lines += "      <c r=`"A$r4`" s=`"2`"><v>$($d.Ano)</v></c>"
    $s4Lines += "      <c r=`"B$r4`" s=`"2`"><v>$($d.SM)</v></c>"
    $s4Lines += "      <c r=`"C$r4`" s=`"2`"><v>$($d.CB)</v></c>"
    $s4Lines += "      <c r=`"D$r4`" s=`"2`"><v>$($d.Horas)</v></c>"
    $s4Lines += "      <c r=`"E$r4`" s=`"2`" t=`"inlineStr`"><is><t>$($d.Perc)</t></is></c>"
    $cleanImp = [System.Security.SecurityElement]::Escape($d.Imp)
    $s4Lines += "      <c r=`"F$r4`" s=`"2`" t=`"inlineStr`"><is><t>$cleanImp</t></is></c>"
    $s4Lines += "    </row>"
}
$s4Lines += '  </sheetData>'
$s4Lines += '</worksheet>'
[System.IO.File]::WriteAllLines("$buildDir\xl\worksheets\sheet4.xml", $s4Lines, [System.Text.Encoding]::UTF8)

# ----------------- SHEET 5: Destino_Impostos_IRBES -----------------
$s5 = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>
    <col min="1" max="1" width="34"/>
    <col min="2" max="2" width="22"/>
    <col min="3" max="3" width="22"/>
    <col min="4" max="4" width="38"/>
  </cols>
  <sheetData>
    <row r="1">
      <c r="A1" s="3" t="inlineStr"><is><t>PARA ONDE VAI O DINHEIRO DOS IMPOSTOS E COMPARATIVO INTERNACIONAL (IRBES)</t></is></c>
    </row>
    <row r="2">
      <c r="A2" s="1" t="inlineStr"><is><t>Destino dos Gastos Publicos Federais</t></is></c>
      <c r="B2" s="1" t="inlineStr"><is><t>% do Orcamento Primario</t></is></c>
      <c r="C2" s="1" t="inlineStr"><is><t>Equivalente no PIB (% PIB)</t></is></c>
      <c r="D2" s="1" t="inlineStr"><is><t>Analise Estrutural e de Eficiencia</t></is></c>
    </row>
    <row r="3">
      <c r="A3" s="2" t="inlineStr"><is><t>Previdencia Social (RGPS + Servidores)</t></is></c>
      <c r="B3" s="2" t="inlineStr"><is><t>44,0% a 46,0%</t></is></c>
      <c r="C3" s="2" t="inlineStr"><is><t>8,2% a 8,8% do PIB</t></is></c>
      <c r="D3" s="2" t="inlineStr"><is><t>Maior rubrica do Estado brasileiro; deficit estrutural crescente</t></is></c>
    </row>
    <row r="4">
      <c r="A4" s="2" t="inlineStr"><is><t>Folha de Pessoal Ativo e Encargos</t></is></c>
      <c r="B4" s="2" t="inlineStr"><is><t>18,0% a 20,0%</t></is></c>
      <c r="C4" s="2" t="inlineStr"><is><t>3,2% a 3,6% do PIB</t></is></c>
      <c r="D4" s="2" t="inlineStr"><is><t>Despesa obrigatoria com estabilidade constitucional</t></is></c>
    </row>
    <row r="5">
      <c r="A5" s="2" t="inlineStr"><is><t>Beneficios Sociais (Bolsa Familia, BPC, Seguro)</t></is></c>
      <c r="B5" s="2" t="inlineStr"><is><t>14,0% a 16,0%</t></is></c>
      <c r="C5" s="2" t="inlineStr"><is><t>2,8% a 3,2% do PIB</t></is></c>
      <c r="D5" s="2" t="inlineStr"><is><t>Transferencias de renda diretas para populacao vulneravel</t></is></c>
    </row>
    <row r="6">
      <c r="A6" s="2" t="inlineStr"><is><t>Saude Publica (SUS - esfera federal)</t></is></c>
      <c r="B6" s="2" t="inlineStr"><is><t>8,5% a 9,5%</t></is></c>
      <c r="C6" s="2" t="inlineStr"><is><t>1,7% a 1,9% do PIB</t></is></c>
      <c r="D6" s="2" t="inlineStr"><is><t>Complementado por estados e municipios; sistema sobrecarregado</t></is></c>
    </row>
    <row r="7">
      <c r="A7" s="2" t="inlineStr"><is><t>Educacao Federal e Repasses (MEC/FNDE)</t></is></c>
      <c r="B7" s="2" t="inlineStr"><is><t>5,5% a 6,5%</t></is></c>
      <c r="C7" s="2" t="inlineStr"><is><t>1,1% a 1,3% do PIB</t></is></c>
      <c r="D7" s="2" t="inlineStr"><is><t>Maior parte do FUNDEB vem de tributos estaduais e municipais</t></is></c>
    </row>
    <row r="8">
      <c r="A8" s="2" t="inlineStr"><is><t>Custeio Administrativo da Maquina Publica</t></is></c>
      <c r="B8" s="2" t="inlineStr"><is><t>7,0% a 9,0%</t></is></c>
      <c r="C8" s="2" t="inlineStr"><is><t>1,3% a 1,6% do PIB</t></is></c>
      <c r="D8" s="2" t="inlineStr"><is><t>Manutencao dos orgaos, frotas, energia, contratos e diárias</t></is></c>
    </row>
    <row r="9">
      <c r="A9" s="2" t="inlineStr"><is><t>Investimentos Publicos em Infraestrutura (Obras)</t></is></c>
      <c r="B9" s="2" t="inlineStr"><is><t>1,5% a 2,5%</t></is></c>
      <c r="C9" s="2" t="inlineStr"><is><t>0,3% a 0,5% do PIB</t></is></c>
      <c r="D9" s="2" t="inlineStr"><is><t>Patamar critico: insuficiente ate para cobrir depreciacao</t></is></c>
    </row>
    <row r="10">
      <c r="A10" s="2" t="inlineStr"><is><t>Juros Nominais da Divida Publica (Financeiro)</t></is></c>
      <c r="B10" s="2" t="inlineStr"><is><t>Extra-primario</t></is></c>
      <c r="C10" s="2" t="inlineStr"><is><t>5,5% a 7,5% do PIB</t></is></c>
      <c r="D10" s="2" t="inlineStr"><is><t>Custo anual da divida bruta de ~78% do PIB com juros elevados</t></is></c>
    </row>
    <row r="12">
      <c r="A12" s="3" t="inlineStr"><is><t>RANKING IRBES (IBPT) - RETORNO DOS TRIBUTOS AO BEM-ESTAR (30 MAIORES CARGAS)</t></is></c>
    </row>
    <row r="13">
      <c r="A13" s="1" t="inlineStr"><is><t>Posicao no Ranking</t></is></c>
      <c r="B13" s="1" t="inlineStr"><is><t>Pais</t></is></c>
      <c r="C13" s="1" t="inlineStr"><is><t>Carga Tributaria (% PIB)</t></is></c>
      <c r="D13" s="1" t="inlineStr"><is><t>IDH e Retorno a Sociedade</t></is></c>
    </row>
    <row r="14">
      <c r="A14" s="2" t="inlineStr"><is><t>1º Lugar</t></is></c>
      <c r="B14" s="2" t="inlineStr"><is><t>Irlanda</t></is></c>
      <c r="C14" s="2" t="inlineStr"><is><t>21,1%</t></is></c>
      <c r="D14" s="2" t="inlineStr"><is><t>IDH 0,950 - Maior eficiencia do mundo</t></is></c>
    </row>
    <row r="15">
      <c r="A15" s="2" t="inlineStr"><is><t>2º Lugar</t></is></c>
      <c r="B15" s="2" t="inlineStr"><is><t>Suica</t></is></c>
      <c r="C15" s="2" t="inlineStr"><is><t>27,9%</t></is></c>
      <c r="D15" s="2" t="inlineStr"><is><t>IDH 0,967 - Servicos publicos e renda excepcionais</t></is></c>
    </row>
    <row r="16">
      <c r="A16" s="2" t="inlineStr"><is><t>3º Lugar</t></is></c>
      <c r="B16" s="2" t="inlineStr"><is><t>Australia</t></is></c>
      <c r="C16" s="2" t="inlineStr"><is><t>28,7%</t></is></c>
      <c r="D16" s="2" t="inlineStr"><is><t>IDH 0,946 - Alta qualidade de vida e educacao de ponta</t></is></c>
    </row>
    <row r="17">
      <c r="A17" s="2" t="inlineStr"><is><t>4º Lugar</t></is></c>
      <c r="B17" s="2" t="inlineStr"><is><t>Estados Unidos</t></is></c>
      <c r="C17" s="2" t="inlineStr"><is><t>27,7%</t></is></c>
      <c r="D17" s="2" t="inlineStr"><is><t>IDH 0,927 - Altissima renda per capita</t></is></c>
    </row>
    <row r="18">
      <c r="A18" s="2" t="inlineStr"><is><t>10º Lugar</t></is></c>
      <c r="B18" s="2" t="inlineStr"><is><t>Noruega</t></is></c>
      <c r="C18" s="2" t="inlineStr"><is><t>42,2%</t></is></c>
      <c r="D18" s="2" t="inlineStr"><is><t>IDH 0,966 - Alta carga, porem servicos de primeiro mundo</t></is></c>
    </row>
    <row r="19">
      <c r="A19" s="2" t="inlineStr"><is><t>20º Lugar</t></is></c>
      <c r="B19" s="2" t="inlineStr"><is><t>Uruguai</t></is></c>
      <c r="C19" s="2" t="inlineStr"><is><t>27,1%</t></is></c>
      <c r="D19" s="2" t="inlineStr"><is><t>IDH 0,830 - Supera o Brasil em qualidade de vida</t></is></c>
    </row>
    <row r="20">
      <c r="A20" s="2" t="inlineStr"><is><t>29º Lugar</t></is></c>
      <c r="B20" s="2" t="inlineStr"><is><t>Argentina</t></is></c>
      <c r="C20" s="2" t="inlineStr"><is><t>29,4%</t></is></c>
      <c r="D20" s="2" t="inlineStr"><is><t>IDH 0,849 - Apesar de crises, entrega IDH maior que o Brasil</t></is></c>
    </row>
    <row r="21">
      <c r="A21" s="2" t="inlineStr"><is><t>30º Lugar (LANTERNA)</t></is></c>
      <c r="B21" s="2" t="inlineStr"><is><t>BRASIL</t></is></c>
      <c r="C21" s="2" t="inlineStr"><is><t>32,4% a 33,7%</t></is></c>
      <c r="D21" s="2" t="inlineStr"><is><t>IDH 0,760 - Ultimo lugar consecutivo em retorno tributario</t></is></c>
    </row>
  </sheetData>
</worksheet>
'@
[System.IO.File]::WriteAllText("$buildDir\xl\worksheets\sheet5.xml", $s5, [System.Text.Encoding]::UTF8)

# Comprime o pacote OpenXML gerando o arquivo .xlsx final
[System.IO.Compression.ZipFile]::CreateFromDirectory($buildDir, $xlsxPath)
Remove-Item -Recurse -Force $buildDir

Write-Host "Arquivo .xlsx gerado com sucesso: $xlsxPath"
Write-Host "Tamanho: $((Get-Item $xlsxPath).Length) bytes"

# Gera tambem a versao CSV consolidada (UTF-8 com BOM e delimitador ponto e virgula)
$csvContent = @"
`"CATEGORIA`";`"ANO_OU_ITEM`";`"CARGA_OU_VALOR`";`"INDICADOR_RELACIONADO`";`"DETALHAMENTO_E_IMPACTO_NA_QUALIDADE_DE_VIDA`"
`"RESUMO_COMPARATIVO`";`"1995-1999 (Menos Imposto)`";`"26,7% a 28,0% do PIB`";`"Dias de trabalho p/ imposto: 102 dias`";`"Inicio do Plano Real; estabilizacao de precos; carga tributaria moderada`"
`"RESUMO_COMPARATIVO`";`"2000-2010 (Expansao Fiscal)`";`"30,0% a 32,8% do PIB`";`"Dias de trabalho p/ imposto: 138 dias`";`"Boom das commodities; aumento de aliquotas de Cofins e PIS; criacao de programas sociais`"
`"RESUMO_COMPARATIVO`";`"2011-2019 (Crise e Plato)`";`"32,1% a 32,5% do PIB`";`"Dias de trabalho p/ imposto: 151 dias`";`"Desoneracoes setoriais fracassadas; recessao historica (2015-16); teto de gastos`"
`"RESUMO_COMPARATIVO`";`"2020-2025 (Mais Imposto / Atual)`";`"32,4% a 33,9% do PIB`";`"Dias de trabalho p/ imposto: 149 dias`";`"Arrecadacao nominal recorde (~R$ 3,9 trilhoes); cidadão trabalha 5 meses so p/ impostos`"
`"EDUCACAO_VS_GASTO`";`"2000`";`"Gasto Educacao: 3,5% PIB`";`"PISA Leitura: 396 pts`";`"Primeiro ano do PISA da OCDE; Brasil na lanterna mundial`"
`"EDUCACAO_VS_GASTO`";`"2005`";`"Gasto Educacao: 3,9% PIB`";`"IDEB Inicial: 3,8 | Medio: 3,4`";`"Criacao oficial do IDEB pelo MEC/Inep`"
`"EDUCACAO_VS_GASTO`";`"2009`";`"Gasto Educacao: 4,8% PIB`";`"PISA Mat: 386 | Lei: 412 | Cie: 405`";`"Melhoria pontual do PISA apos universalizacao do ensino fundamental`"
`"EDUCACAO_VS_GASTO`";`"2015`";`"Gasto Educacao: 5,5% PIB`";`"IDEB Medio: 3,7 | PISA Mat: 377`";`"Gasto cresce fortemente, mas resultados de aprendizagem entram em estagnacao`"
`"EDUCACAO_VS_GASTO`";`"2022`";`"Gasto Educacao: 5,6% PIB`";`"PISA Mat: 379 | Lei: 410 | Cie: 403`";`"Brasil no 64º lugar de 81 paises; mais de 70% dos jovens sem base em matematica`"
`"EDUCACAO_VS_GASTO`";`"2023`";`"Gasto Educacao: 5,7% PIB`";`"IDEB Inicial: 6,0 | Medio: 4,3`";`"Anos iniciais bateu meta nacional (6,0), mas Ensino Medio segue em crise profunda`"
`"PODER_DE_COMPRA`";`"1995`";`"Salario Minimo: R$ 100`";`"Cesta Basica: R$ 65,10`";`"Trabalhador precisava de 143 horas de trabalho p/ comprar a cesta (65,1% do piso)`"
`"PODER_DE_COMPRA`";`"2005`";`"Salario Minimo: R$ 300`";`"Cesta Basica: R$ 166,50`";`"Trabalhador precisava de 122 horas de trabalho (55,5% do piso)`"
`"PODER_DE_COMPRA`";`"2014`";`"Salario Minimo: R$ 724`";`"Cesta Basica: R$ 313,20`";`"Pico do poder de compra: 95 horas trabalhadas (43,3% do piso consumido)`"
`"PODER_DE_COMPRA`";`"2021`";`"Salario Minimo: R$ 1.100`";`"Cesta Basica: R$ 692,30`";`"Inflacao de alimentos eleva esforco para 122 horas (62,9% do piso consumido)`"
`"PODER_DE_COMPRA`";`"2025`";`"Salario Minimo: R$ 1.518`";`"Cesta Basica: R$ 845,00`";`"Necessarias 107 horas de trabalho; impostos no consumo (ICMS/PIS/Cofins) pesam sobre comida`"
`"DESTINO_DOS_IMPOSTOS`";`"Previdencia Social`";`"8,5% do PIB (45% do orcamento federal)`";`"Gastos com aposentadorias e pensoes`";`"Maior rigidez do Estado; pressao demografica sem aumento proporcional de servicos`"
`"DESTINO_DOS_IMPOSTOS`";`"Pessoal e Folha`";`"3,4% do PIB (19% do orcamento)`";`"Salarios de servidores e judiciario`";`"Gasto de carater fixo e obrigatorio`"
`"DESTINO_DOS_IMPOSTOS`";`"Investimento Publico`";`"0,4% do PIB (2% do orcamento)`";`"Obras, saneamento, rodovias, ferrovias`";`"Queda historica de mais de 70%; infraestrutura sucateada reduz a produtividade do pais`"
`"DESTINO_DOS_IMPOSTOS`";`"Juros da Divida`";`"6,5% do PIB (extra-primario)`";`"Rolagem e servico da divida publica`";`"Elevado endividamento e Selic drenam recursos que poderiam ir para servicos essenciais`"
`"RANKING_IRBES_IBPT`";`"1º Lugar: Irlanda`";`"Carga: 21,1% do PIB`";`"IDH: 0,950`";`"Maior retorno de qualidade de vida do mundo por imposto pago`"
`"RANKING_IRBES_IBPT`";`"2º Lugar: Suica`";`"Carga: 27,9% do PIB`";`"IDH: 0,967`";`"Alta renda, saude e educacao de primeiro mundo com carga menor que a brasileira`"
`"RANKING_IRBES_IBPT`";`"20º Lugar: Uruguai`";`"Carga: 27,1% do PIB`";`"IDH: 0,830`";`"Pais sul-americano que entrega muito mais retorno social que o Brasil`"
`"RANKING_IRBES_IBPT`";`"30º Lugar: BRASIL`";`"Carga: 32,4% a 33,7% do PIB`";`"IDH: 0,760`";`"LANTERNA: Menor retorno em bem-estar social entre os 30 paises de maior arrecadacao`"
"@

[System.IO.File]::WriteAllText($csvPath, $csvContent, [System.Text.Encoding]::UTF8)
Write-Host "Arquivo CSV gerado com sucesso: $csvPath"
Write-Host "Tamanho CSV: $((Get-Item $csvPath).Length) bytes"
