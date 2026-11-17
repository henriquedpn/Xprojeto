// Verificação de tipos das fórmulas geradas, simulando a lista SharePoint e os controles.
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.PowerFx;
using Microsoft.PowerFx.Types;

var formulas = JsonDocument.Parse(File.ReadAllText(args[0])).RootElement;

var S = FormulaType.String;
var pessoa = RecordType.Empty().Add("Claims", S).Add("DisplayName", S).Add("Email", S)
    .Add("Department", S).Add("JobTitle", S).Add("Picture", S);
var escolha = RecordType.Empty().Add("Value", S);
var projeto = RecordType.Empty()
    .Add("ID", FormulaType.Number).Add("Title", S).Add("Descricao", S)
    .Add("Colaborador", pessoa).Add("Gestor", pessoa).Add("Disciplina", escolha)
    .Add("DataInicio", FormulaType.Date).Add("DuracaoDias", FormulaType.Number).Add("DataFim", FormulaType.Date)
    .Add("StatusProjeto", escolha).Add("PercConcluido", FormulaType.Number).Add("Observacoes", S);
var tabProjetos = projeto.ToTable();
var pessoaResumo = RecordType.Empty().Add("Email", S).Add("Nome", S).Add("Primeiro", FormulaType.Boolean);

var config = new PowerFxConfig(Features.PowerFxV1);
config.EnableSetFunction();
config.SymbolTable.EnableMutationFunctions();
config.AddFunction(new NotifyFn());
config.AddFunction(new NavigateFn());

var st = new SymbolTable();
st.AddVariable("Projetos", tabProjetos, mutable: true);
st.AddVariable("colProj", tabProjetos, mutable: true);
st.AddVariable("colPessoas", pessoaResumo.ToTable(), mutable: true);
st.AddVariable("gItem", projeto, mutable: true);
st.AddVariable("gSalvo", projeto, mutable: true);
st.AddVariable("gEmail", S, mutable: true);
st.AddVariable("locConfirmar", FormulaType.Boolean, mutable: true);
st.AddVariable("locPessoa", S, mutable: true);
st.AddVariable("defaultsProjetos", projeto);
st.AddVariable("errosProjetos", RecordType.Empty().Add("Message", S).ToTable());
st.AddVariable("usuario", RecordType.Empty().Add("Email", S).Add("FullName", S));
st.AddVariable("choicesPessoa", pessoa.ToTable());
st.AddVariable("choicesEscolha", escolha.ToTable());
foreach (var scr in new[] { "scrProjetos", "scrFormulario", "scrDashboard" }) st.AddVariable(scr, S);
RecordType Enum(params string[] campos) { var r = RecordType.Empty(); foreach (var c in campos) r = r.Add(c, S); return r; }
st.AddVariable("ScreenTransition", Enum("None", "Fade"));
st.AddVariable("NotificationType", Enum("Error", "Success", "Information"));
st.AddVariable("TextMode", Enum("MultiLine", "SingleLine"));
st.AddVariable("TextFormat", Enum("Number", "Text"));
st.AddVariable("Align", Enum("Left", "Center", "Right"));
st.AddVariable("FontWeight", Enum("Semibold", "Bold", "Normal"));
// controles referenciados
st.AddVariable("txtCodigo", RecordType.Empty().Add("Text", S));
st.AddVariable("txtDescricao", RecordType.Empty().Add("Text", S));
st.AddVariable("txtDuracao", RecordType.Empty().Add("Text", S));
st.AddVariable("txtObs", RecordType.Empty().Add("Text", S));
st.AddVariable("txtBusca", RecordType.Empty().Add("Text", S));
st.AddVariable("cmbColaborador", RecordType.Empty().Add("Selected", pessoa));
st.AddVariable("cmbGestor", RecordType.Empty().Add("Selected", pessoa));
st.AddVariable("drpDisciplina", RecordType.Empty().Add("Selected", escolha));
st.AddVariable("drpStatus", RecordType.Empty().Add("Selected", escolha));
st.AddVariable("drpVisao", RecordType.Empty().Add("Selected", escolha));
st.AddVariable("dteInicio", RecordType.Empty().Add("SelectedDate", FormulaType.Date));
st.AddVariable("sldPerc", RecordType.Empty().Add("Value", FormulaType.Number));
st.AddVariable("galProjetos", RecordType.Empty().Add("AllItems", tabProjetos));
st.AddVariable("galDetalhe", RecordType.Empty().Add("AllItems", tabProjetos));
foreach (var b in new[] { "Ini", "And", "Pau", "Atr", "Con" })
    st.AddVariable("recBar" + b, RecordType.Empty().Add("X", FormulaType.Number).Add("Width", FormulaType.Number));
foreach (var p in new[] { "Width", "Height", "TemplateWidth", "TemplateHeight" }) st.AddVariable("pai" + p, FormulaType.Number);

var escopos = new Dictionary<string, ReadOnlySymbolTable>();
SymbolTable Item(RecordType t) { var s = new SymbolTable(); s.AddVariable("ThisItem", t); return s; }
escopos["galProjetos"] = ReadOnlySymbolTable.Compose(Item(projeto), st);
escopos["galDetalhe"] = ReadOnlySymbolTable.Compose(Item(projeto), st);
escopos["galResumo"] = ReadOnlySymbolTable.Compose(Item(pessoaResumo), st);

var engine = new RecalcEngine(config);
var opts = new ParserOptions { AllowsSideEffects = true, Culture = new System.Globalization.CultureInfo("en-US") };

// Funções exclusivas do Power Apps -> equivalentes verificáveis
string Adaptar(string f)
{
    f = Regex.Replace(f, @"\bChoices\(Projetos\.(Colaborador|Gestor)\)", "choicesPessoa");
    f = Regex.Replace(f, @"\bChoices\(Projetos\.\w+\)", "choicesEscolha");
    f = f.Replace("Errors(Projetos)", "errosProjetos").Replace("Defaults(Projetos)", "defaultsProjetos")
         .Replace("User()", "usuario").Replace("Refresh(Projetos)", "true")
         .Replace("Select(Parent)", "true").Replace("Select(btnDashAtualizar)", "true")
         .Replace("ClearCollect(", "Set(");
    f = Regex.Replace(f, @"\bReset\(\w+\)", "true");
    f = Regex.Replace(f, @"\bParent\.(Width|Height|TemplateWidth|TemplateHeight)\b", "pai$1");
    f = Regex.Replace(f, @"UpdateContext\(\{(\w+): (.*)\}\)", "Set($1, $2)");
    f = Regex.Replace(f, @"(?<![.\w])Days\)", "TimeUnit.Days)");
    return f;
}

var numericas = new HashSet<string> { "X", "Y", "Width", "Height", "Size", "ZIndex", "TemplateSize", "TemplatePadding",
    "BorderThickness", "RadiusTopLeft", "RadiusTopRight", "RadiusBottomLeft", "RadiusBottomRight", "Min", "Max" };
var cores = new HashSet<string> { "Fill", "Color", "HoverFill", "PressedFill" };
var logicas = new HashSet<string> { "Visible", "ShowScrollbar", "ShowValue", "SelectMultiple", "IsSearchable",
    "AllowEmptySelection", "BackEnabled" };
var ignorar = new HashSet<string> { "FontWeight", "Align", "Mode", "Format" };   // enums do Power Apps

int total = 0, erros = 0;
foreach (var f in formulas.EnumerateArray())
{
    string ctl = f.GetProperty("controle").GetString()!, prop = f.GetProperty("prop").GetString()!;
    string modelo = f.GetProperty("modelo").GetString()!;
    if (ignorar.Contains(prop)) continue;
    var esc = f.GetProperty("escopo").ValueKind == JsonValueKind.String ? f.GetProperty("escopo").GetString()! : null;
    var texto = Adaptar(f.GetProperty("formula").GetString()!);
    var res = engine.Check(texto, opts, esc != null ? escopos[esc] : st);
    total++;
    var msgs = res.Errors.Where(e => !e.IsWarning).Select(e => e.Message).ToList();
    if (res.IsSuccess)
    {
        var t = res.ReturnType;
        if (numericas.Contains(prop) && t is not NumberType && t is not DecimalType) msgs.Add($"esperado número, veio {t}");
        if (cores.Contains(prop) && t is not ColorType) msgs.Add($"esperado cor, veio {t}");
        if (logicas.Contains(prop) && t is not BooleanType) msgs.Add($"esperado booleano, veio {t}");
        if (prop is "Text" or "HintText" && t is not StringType && t is not NumberType && t is not DecimalType) msgs.Add($"esperado texto, veio {t}");
        if (prop == "Items" && t is not TableType) msgs.Add($"esperado tabela, veio {t}");
        if (prop == "DefaultDate" && t is not DateType) msgs.Add($"esperado data, veio {t}");
        if (prop == "Default" && modelo is "text" or "dropdown" && t is not StringType) msgs.Add($"esperado texto, veio {t}");
        if (prop == "Default" && modelo == "slider" && t is not NumberType && t is not DecimalType) msgs.Add($"esperado número, veio {t}");
    }
    if (msgs.Count > 0)
    {
        erros++;
        Console.WriteLine($"ERRO {ctl}.{prop}: {string.Join(" | ", msgs.Distinct())}");
    }
}
Console.WriteLine($"Fórmulas verificadas: {total}; com erro: {erros}");

class NotifyFn : ReflectionFunction
{
    public NotifyFn() : base("Notify", FormulaType.Boolean, FormulaType.String, FormulaType.String) { }
    public BooleanValue Execute(StringValue a, StringValue b) => FormulaValue.New(true);
}
class NavigateFn : ReflectionFunction
{
    public NavigateFn() : base("Navigate", FormulaType.Boolean, FormulaType.String, FormulaType.String) { }
    public BooleanValue Execute(StringValue a, StringValue b) => FormulaValue.New(true);
}
