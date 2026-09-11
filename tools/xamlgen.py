"""
xamlgen.py
==========
Emits UiPath Windows (.NET 8) workflow XAML for Studio 25.10.

The namespace and assembly-reference blocks are lifted from Studio's own
`ProjectTemplates/Blank/Windows/VisualBasic/Main.xaml`, so generated files load in the
designer exactly like hand-drawn ones. Studio rewrites them on first save; nothing here
survives as a special format.

Two things this module exists to get right, because both cost a debugging cycle each
time they are done by hand:

1. **Attribute escaping.** VB source and VB expressions go into XAML *attributes*, so
   newlines must become `&#xA;` and `<`, `&`, `"` must be entity-escaped. A property
   element with `xml:space="preserve"` is NOT a legal alternative - the XAML parser
   rejects an attribute on a non-empty property element.
2. **Class names.** `x:Class` becomes a VB class, so it cannot start with a digit.
   Numbered file names get an underscore prefix on the class only.
"""

from __future__ import annotations

import re

# --------------------------------------------------------------------------
# Namespaces
# --------------------------------------------------------------------------
BASE_NAMESPACES = {
    "": "http://schemas.microsoft.com/netfx/2009/xaml/activities",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "mva": "clr-namespace:Microsoft.VisualBasic.Activities;assembly=System.Activities",
    "njl": "clr-namespace:Newtonsoft.Json.Linq;assembly=Newtonsoft.Json",
    "sap": "http://schemas.microsoft.com/netfx/2009/xaml/activities/presentation",
    "sap2010": "http://schemas.microsoft.com/netfx/2010/xaml/activities/presentation",
    "scg": "clr-namespace:System.Collections.Generic;assembly=System.Private.CoreLib",
    "sco": "clr-namespace:System.Collections.ObjectModel;assembly=mscorlib",
    "s": "clr-namespace:System;assembly=System.Private.CoreLib",
    "sd": "clr-namespace:System.Data;assembly=System.Data.Common",
    "ui": "http://schemas.uipath.com/workflow/activities",
    "x": "http://schemas.microsoft.com/winfx/2006/xaml",
}

STD_IMPORTS = [
    "System.Activities", "System.Activities.Statements", "System.Activities.Expressions",
    "System.Activities.Validation", "System.Activities.XamlIntegration",
    "Microsoft.VisualBasic", "Microsoft.VisualBasic.Activities",
    "System", "System.Collections", "System.Collections.Generic",
    "System.Collections.ObjectModel", "System.Data", "System.Diagnostics",
    "System.Drawing", "System.Globalization", "System.IO", "System.Linq",
    "System.Net.Mail", "System.Security.Cryptography", "System.Text",
    "System.Xml", "System.Xml.Linq",
    "Newtonsoft.Json", "Newtonsoft.Json.Linq",
    "UiPath.Core", "UiPath.Core.Activities", "System.Windows.Markup",
]

STD_REFS = [
    "System.Activities", "Microsoft.VisualBasic", "System.Private.CoreLib", "mscorlib",
    "System.Data", "System", "System.Linq", "System.Data.Common", "System.Drawing",
    "System.Drawing.Primitives", "System.Drawing.Common", "System.Core", "System.Runtime",
    "System.Security.Cryptography", "System.Xml", "System.Xml.Linq",
    "PresentationFramework", "WindowsBase", "PresentationCore", "System.Xaml",
    "Newtonsoft.Json",
    "UiPath.System.Activities", "UiPath.UiAutomation.Activities",
    "UiPath.Excel.Activities", "UiPath.Excel", "UiPath.Mail.Activities",
    # HttpClient and DeserializeJson live here, not in UiPath.System.Activities.
    # The xmlns -> type resolution walks this list, so a missing reference shows up
    # as "Cannot create unknown type ...DeserializeJson" rather than a missing package.
    "UiPath.Web.Activities",
    "System.Data.DataSetExtensions", "System.Net.Mail",
]


# --------------------------------------------------------------------------
# Escaping
# --------------------------------------------------------------------------
def esc(value: str) -> str:
    """Escape a string for use inside a XAML attribute, preserving newlines."""
    return (value
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
            .replace("\r\n", "\n")
            .replace("\r", "\n")
            .replace("\n", "&#xA;")
            .replace("\t", "&#x9;"))


def vb(expression: str) -> str:
    """Wrap a VB expression in the [ ] UiPath uses, escaped for an attribute."""
    return esc(f"[{expression}]")


def lit(text: str) -> str:
    """A plain literal attribute value (not a VB expression)."""
    return esc(text)


def attrs(**kwargs) -> str:
    """Render attributes, skipping None. Values must already be escaped."""
    return "".join(f' {k.rstrip("_").replace("__", ":")}="{v}"'
                   for k, v in kwargs.items() if v is not None)


# --------------------------------------------------------------------------
# Activity helpers
# --------------------------------------------------------------------------
def log(message_expr: str, level: str = "Info", name: str | None = None) -> str:
    """
    LogMessage. `message_expr` is raw VB - wrapped and escaped here.

    When no name is given, one is derived from the message's leading literal so every
    activity gets a distinct, meaningful display name. Leaving them all as the default
    "Log Message" trips the Workflow Analyzer's ST-MRD-002 and ST-NMG-004 rules and
    makes the designer much harder to read.
    """
    if name is None:
        m = re.search(r'"([^"]{2,60})', message_expr)
        label = m.group(1).strip() if m else "message"
        label = re.sub(r"\s+", " ", label).strip(" -:&")
        name = f"Log - {label[:46]}".rstrip()
    return f'<ui:LogMessage DisplayName="{lit(name)}" Level="{level}" Message="{vb(message_expr)}" />'


def assign(to_var: str, value_expr: str, type_ref: str = "x:String",
           name: str | None = None) -> str:
    label = name or f"Assign - {to_var}"
    return (f'<Assign DisplayName="{lit(label)}">'
            f'<Assign.To><OutArgument x:TypeArguments="{type_ref}">{vb(to_var)}</OutArgument></Assign.To>'
            f'<Assign.Value><InArgument x:TypeArguments="{type_ref}">{vb(value_expr)}</InArgument></Assign.Value>'
            f'</Assign>')


def variables(*specs: tuple[str, str]) -> str:
    """specs are (type_ref, name), e.g. ("sd:DataTable", "dtChain")."""
    if not specs:
        return ""
    inner = "".join(f'<Variable x:TypeArguments="{t}" Name="{n}" />' for t, n in specs)
    return f"<Sequence.Variables>{inner}</Sequence.Variables>"


def sequence(display: str, body: str, vars_xml: str = "") -> str:
    return f'<Sequence DisplayName="{lit(display)}">{vars_xml}{body}</Sequence>'


def invoke_code(code: str, arguments: list[tuple[str, str, str, str]],
                name: str = "Invoke Code") -> str:
    """
    arguments: list of (direction, key, type_ref, expression) where direction is
    'In', 'Out' or 'InOut'. `expression` is raw VB (a variable name, usually).
    """
    args_xml = "".join(
        f'<{d}Argument x:TypeArguments="{t}" x:Key="{k}">{vb(e)}</{d}Argument>'
        for d, k, t, e in arguments)
    return (f'<ui:InvokeCode DisplayName="{lit(name)}" Language="VBNet" Code="{esc(code)}">'
            f'<ui:InvokeCode.Arguments>'
            f'<scg:Dictionary x:TypeArguments="x:String, Argument">{args_xml}</scg:Dictionary>'
            f'</ui:InvokeCode.Arguments></ui:InvokeCode>')


def invoke_workflow(path: str, arguments: list[tuple[str, str, str, str]],
                    name: str | None = None) -> str:
    """Invoke Workflow File. `path` is relative to the project root."""
    label = name or f"Invoke {path}"
    args_xml = "".join(
        f'<{d}Argument x:TypeArguments="{t}" x:Key="{k}">{vb(e)}</{d}Argument>'
        for d, k, t, e in arguments)
    args_block = (f'<ui:InvokeWorkflowFile.Arguments>'
                  f'<scg:Dictionary x:TypeArguments="x:String, Argument">{args_xml}</scg:Dictionary>'
                  f'</ui:InvokeWorkflowFile.Arguments>') if arguments else ""
    # A plain literal, not a VB expression: written as ["..."] the analyzer reports
    # SY-USG-015 and Studio cannot resolve the target at design time, so a typo in a
    # workflow path would only surface at run time.
    return (f'<ui:InvokeWorkflowFile DisplayName="{lit(label)}" '
            f'WorkflowFileName="{lit(path)}">'
            f'{args_block}</ui:InvokeWorkflowFile>')


def switch(expression_expr: str, cases: list[tuple[str, str]], default_xml: str,
           name: str = "Switch") -> str:
    """
    String Switch. The case key goes as `x:Key` on the case *activity* itself -
    wrapping it in an <x:String> element instead fails with "Type 'System.String'
    does not have a content property".
    """
    cases_xml = "".join(
        activity_xml.replace("<Sequence ", f'<Sequence x:Key="{lit(key)}" ', 1)
        for key, activity_xml in cases)
    return (f'<Switch x:TypeArguments="x:String" DisplayName="{lit(name)}" '
            f'Expression="{vb(expression_expr)}">'
            f'<Switch.Default>{default_xml}</Switch.Default>{cases_xml}</Switch>')


def if_(condition_expr: str, then_xml: str, else_xml: str = "", name: str = "If") -> str:
    else_block = f"<If.Else>{else_xml}</If.Else>" if else_xml else ""
    return (f'<If DisplayName="{lit(name)}" Condition="{vb(condition_expr)}">'
            f'<If.Then>{then_xml}</If.Then>{else_block}</If>')


def for_each_row(dt_expr: str, body_xml: str, row_var: str = "CurrentRow",
                 name: str = "For Each Row in DataTable") -> str:
    return (f'<ui:ForEachRow DisplayName="{lit(name)}" DataTable="{vb(dt_expr)}">'
            f'<ui:ForEachRow.Body>'
            f'<ActivityAction x:TypeArguments="sd:DataRow">'
            f'<ActivityAction.Argument>'
            f'<DelegateInArgument x:TypeArguments="sd:DataRow" Name="{row_var}" />'
            f'</ActivityAction.Argument>{body_xml}</ActivityAction>'
            f'</ui:ForEachRow.Body></ui:ForEachRow>')


def excel_scope(path_expr: str, body_xml: str, name: str = "Excel Application Scope",
                autosave: bool = False, visible: bool = False,
                create_new: bool = False, read_only: bool = True) -> str:
    """
    Classic Excel Application Scope. The body's delegate argument MUST be named
    `ExcelWorkbookScope` - child activities resolve the workbook out of the data
    context by that exact name, and a different name fails only at runtime.
    """
    return (f'<ui:ExcelApplicationScope DisplayName="{lit(name)}" '
            f'AutoSave="{str(autosave)}" Visible="{str(visible)}" '
            f'CreateNewFile="{str(create_new)}" ReadOnly="{str(read_only)}" '
            f'WorkbookPath="{vb(path_expr)}">'
            f'<ui:ExcelApplicationScope.Body>'
            f'<ActivityAction x:TypeArguments="ui:WorkbookApplication">'
            f'<ActivityAction.Argument>'
            f'<DelegateInArgument x:TypeArguments="ui:WorkbookApplication" Name="ExcelWorkbookScope" />'
            f'</ActivityAction.Argument>{body_xml}</ActivityAction>'
            f'</ui:ExcelApplicationScope.Body></ui:ExcelApplicationScope>')


def excel_read(sheet: str, into_var: str, name: str | None = None,
               add_headers: bool = True, rng: str = "") -> str:
    label = name or f"Read Range - {sheet}"
    return (f'<ui:ExcelReadRange DisplayName="{lit(label)}" AddHeaders="{str(add_headers)}" '
            f'DataTable="{vb(into_var)}" Range="{lit(rng)}" SheetName="{lit(sheet)}" />')


def excel_write(sheet: str, dt_expr: str, name: str | None = None,
                add_headers: bool = True, start_cell: str = "A1") -> str:
    label = name or f"Write Range - {sheet}"
    return (f'<ui:ExcelWriteRange DisplayName="{lit(label)}" AddHeaders="{str(add_headers)}" '
            f'DataTable="{vb(dt_expr)}" SheetName="{lit(sheet)}" StartingCell="{lit(start_cell)}" />')


def read_text(path_expr: str, into_var: str, name: str = "Read Text File") -> str:
    return (f'<ui:ReadTextFile DisplayName="{lit(name)}" '
            f'FileName="{vb(path_expr)}" Content="{vb(into_var)}" />')


def write_text(path_expr: str, content_expr: str, name: str = "Write Text File") -> str:
    return (f'<ui:WriteTextFile DisplayName="{lit(name)}" '
            f'FileName="{vb(path_expr)}" Text="{vb(content_expr)}" />')


def deserialize_json(json_expr: str, into_var: str, name: str = "Deserialize JSON",
                     type_ref: str = "njl:JObject") -> str:
    """
    Deserialize JSON. The activity is generic (`DeserializeJson`1`), so the
    x:TypeArguments is mandatory - without it the type simply does not resolve.
    """
    return (f'<ui:DeserializeJson x:TypeArguments="{type_ref}" DisplayName="{lit(name)}" '
            f'JsonString="{vb(json_expr)}" JsonObject="{vb(into_var)}" />')


def try_catch(try_xml: str, catch_xml: str, exception_type: str = "s:Exception",
              finally_xml: str = "", name: str = "Try Catch",
              ex_var: str = "exception") -> str:
    fin = f"<TryCatch.Finally>{finally_xml}</TryCatch.Finally>" if finally_xml else ""
    return (f'<TryCatch DisplayName="{lit(name)}">'
            f'<TryCatch.Try>{try_xml}</TryCatch.Try>'
            f'<TryCatch.Catches>'
            f'<Catch x:TypeArguments="{exception_type}">'
            f'<ActivityAction x:TypeArguments="{exception_type}">'
            f'<ActivityAction.Argument>'
            f'<DelegateInArgument x:TypeArguments="{exception_type}" Name="{ex_var}" />'
            f'</ActivityAction.Argument>{catch_xml}</ActivityAction>'
            f'</Catch></TryCatch.Catches>{fin}</TryCatch>')


def throw(expr: str, name: str = "Throw") -> str:
    return (f'<Throw DisplayName="{lit(name)}" Exception="{vb(expr)}" />')


# --------------------------------------------------------------------------
# Workflow document
# --------------------------------------------------------------------------
def _class_name(stem: str) -> str:
    """x:Class becomes a VB class name, so it must not start with a digit."""
    return f"_{stem}" if stem and stem[0].isdigit() else stem


def workflow(stem: str, body: str, members: list[tuple[str, str]] | None = None,
             extra_namespaces: dict[str, str] | None = None,
             extra_imports: list[str] | None = None,
             extra_refs: list[str] | None = None) -> str:
    """
    Build a complete workflow document.

    stem     : file name without extension, e.g. "05_Validate_Engine"
    body     : XML for the root activity (normally one <Sequence>)
    members  : arguments as (name, type) e.g. ("in_Config", "InArgument(scg:Dictionary(x:String, x:Object))")
    """
    ns = dict(BASE_NAMESPACES)
    if extra_namespaces:
        ns.update(extra_namespaces)

    ns_attrs = "\n ".join(
        f'xmlns{"" if p == "" else ":" + p}="{u}"' for p, u in sorted(ns.items()))

    imports = STD_IMPORTS + list(extra_imports or [])
    refs = STD_REFS + list(extra_refs or [])

    imports_xml = "\n      ".join(f"<x:String>{i}</x:String>" for i in imports)
    refs_xml = "\n      ".join(f"<AssemblyReference>{r}</AssemblyReference>" for r in refs)

    members_xml = ""
    if members:
        rows = "\n    ".join(
            f'<x:Property Name="{n}" Type="{t}" />' for n, t in members)
        members_xml = f"\n  <x:Members>\n    {rows}\n  </x:Members>\n"

    return f"""<Activity mc:Ignorable="sap sap2010" x:Class="{_class_name(stem)}" mva:VisualBasic.Settings="{{x:Null}}"
 {ns_attrs}>{members_xml}
  <TextExpression.NamespacesForImplementation>
    <sco:Collection x:TypeArguments="x:String">
      {imports_xml}
    </sco:Collection>
  </TextExpression.NamespacesForImplementation>
  <TextExpression.ReferencesForImplementation>
    <sco:Collection x:TypeArguments="AssemblyReference">
      {refs_xml}
    </sco:Collection>
  </TextExpression.ReferencesForImplementation>
  {body}
</Activity>
"""
