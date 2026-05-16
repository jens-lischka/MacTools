#!/usr/bin/env python3
"""Generate the StudioTools ribbon tab into both manifests and commands.ts.

Single source of truth: GROUPS below. Run from the repo root:  python3 scripts/gen-ribbon.py
"""
import os

# (group id suffix, group label, [(button id, label, icon, ts expression)])
GROUPS = [
    ("Align", "Align", [
        ("alignLeft", "Align Left", "AlignLeft", 'align("left", "selection")'),
        ("alignCenter", "Align Center", "AlignCenter", 'align("centerH", "selection")'),
        ("alignRight", "Align Right", "AlignRight", 'align("right", "selection")'),
        ("alignTop", "Align Top", "AlignTop", 'align("top", "selection")'),
        ("alignMiddle", "Align Middle", "AlignMiddle", 'align("middle", "selection")'),
        ("alignBottom", "Align Bottom", "AlignBottom", 'align("bottom", "selection")'),
        ("distributeH", "Distribute Horizontally", "DistributeHorizontally", 'distribute("horizontal")'),
        ("distributeV", "Distribute Vertically", "DistributeVertically", 'distribute("vertical")'),
    ]),
    ("Size", "Size & Position", [
        ("matchWidth", "Match Width", "MatchWidth", 'matchSize("width", "last")'),
        ("matchHeight", "Match Height", "MatchHeight", 'matchSize("height", "last")'),
        ("matchSizeBoth", "Match Size", "MatchSize", 'matchSize("both", "last")'),
        ("stretchLeft", "Stretch Left", "StretchLeft", 'stretchToEdge("left", "selection")'),
        ("stretchRight", "Stretch Right", "StretchRight", 'stretchToEdge("right", "selection")'),
        ("stretchTop", "Stretch Top", "StretchTop", 'stretchToEdge("top", "selection")'),
        ("stretchBottom", "Stretch Bottom", "StretchBottom", 'stretchToEdge("bottom", "selection")'),
        ("fillGapH", "Fill Horizontal Gaps", "FillHorizontalGap", 'fillGap("horizontal")'),
        ("fillGapV", "Fill Vertical Gaps", "FillVerticalGap", 'fillGap("vertical")'),
        ("unifySize", "Unify Size", "UnifyCorners", 'unifyShapes()'),
        ("straightenLines", "Straighten Lines", "StraightenLine", 'straightenLines()'),
    ]),
    ("Select", "Select Same", [
        ("selectFill", "Select Same Fill", "SelectSameFill", 'selectSame("fillColor")'),
        ("selectOutline", "Select Same Outline", "SelectSameOutline", 'selectSame("lineColor")'),
        ("selectWeight", "Select Same Outline Weight", "SelectSameOutlineWeight", 'selectSame("lineWeight")'),
        ("selectFont", "Select Same Font", "SelectSameFontName", 'selectSame("fontName")'),
        ("selectType", "Select Same Type", "SelectSameType", 'selectSame("shapeType")'),
        ("selectSize", "Select Same Size", "SelectSameSize", 'selectSame("size")'),
        ("selectPosTop", "Select Same Top Edge", "SelectSamePositionTop", 'selectSame("positionTop")'),
        ("selectPosLeft", "Select Same Left Edge", "SelectSamePositionLeft", 'selectSame("positionLeft")'),
        ("selectPosRight", "Select Same Right Edge", "SelectSamePositionRight", 'selectSame("positionRight")'),
        ("selectPosBottom", "Select Same Bottom Edge", "SelectSamePositionBottom", 'selectSame("positionBottom")'),
        ("hideShapes", "Hide Selected", "HideObject", "setSelectedShapesVisible(false)"),
        ("showShapes", "Show All Shapes", "ShowAll", "showAllShapes()"),
    ]),
    ("Text", "Text", [
        ("autofitOn", "Resize Shape to Fit Text", "ResizeShapeToFitTextOn", "setAutoSize(PowerPoint.ShapeAutoSize.autoSizeShapeToFitText)"),
        ("autofitOff", "Turn Autofit Off", "ResizeShapeToFitTextMix", "setAutoSize(PowerPoint.ShapeAutoSize.autoSizeNone)"),
        ("wrapOn", "Word Wrap On", "WrapTextOn", "setWordWrap(true)"),
        ("wrapOff", "Word Wrap Off", "WrapTextMix", "setWordWrap(false)"),
        ("marginsNone", "No Margins", "SetMarginsNone", "setMargins({ left: 0, right: 0, top: 0, bottom: 0 })"),
        ("marginsNarrow", "Narrow Margins", "SetMarginsNarrow", "setMargins({ left: 3.6, right: 3.6, top: 3.6, bottom: 3.6 })"),
        ("marginsNormal", "Normal Margins", "SetMarginsNormal", "setMargins({ left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 })"),
        ("marginsWide", "Wide Margins", "SetMarginsWide", "setMargins({ left: 14.4, right: 14.4, top: 14.4, bottom: 14.4 })"),
        ("alignTextLeft", "Align Text Left", "AlignLeft", "setParagraphAlignment(PowerPoint.ParagraphHorizontalAlignment.left)"),
        ("alignTextCenter", "Align Text Center", "AlignCenter", "setParagraphAlignment(PowerPoint.ParagraphHorizontalAlignment.center)"),
        ("alignTextRight", "Align Text Right", "AlignRight", "setParagraphAlignment(PowerPoint.ParagraphHorizontalAlignment.right)"),
        ("alignTextJustify", "Justify Text", "AlignInGrid", "setParagraphAlignment(PowerPoint.ParagraphHorizontalAlignment.justify)"),
        ("clearBreaks", "Clear Line Breaks", "ClearLineBreaks", "clearLineBreaks()"),
        ("clearTextRibbon", "Clear Text", "DeleteText", "clearText()"),
        ("bulletsOn", "Bullets On", "FixBullets", "setBullets(true)"),
        ("bulletsOff", "Bullets Off", "NoBullets", "setBullets(false)"),
        ("styleHeading", "Heading Style", "Heading1Text", "applyTextStyle({ size: 28, bold: true })"),
        ("styleSubheading", "Subheading Style", "Subheading1Text", "applyTextStyle({ size: 20, bold: true })"),
        ("styleBody", "Body Style", "Body1Text", "applyTextStyle({ size: 14, bold: false })"),
        ("mergeTextRibbon", "Merge Text", "MergeText", "mergeText()"),
        ("splitTextRibbon", "Split Text", "SplitText", "splitText()"),
    ]),
    ("Shapes", "Shapes", [
        ("insertRectangle", "Insert Rectangle", "ShapesRectangle", "insertShape(PowerPoint.GeometricShapeType.rectangle)"),
        ("insertRoundRect", "Insert Rounded Rectangle", "ShapesRoundedRectangle", "insertShape(PowerPoint.GeometricShapeType.roundRectangle)"),
        ("insertOval", "Insert Oval", "ShapesOval", "insertShape(PowerPoint.GeometricShapeType.ellipse)"),
        ("insertTriangle", "Insert Triangle", "ShapesLargeCaret", "insertShape(PowerPoint.GeometricShapeType.triangle)"),
        ("insertArrow", "Insert Right Arrow", "SymbolsArrowRight", "insertShape(PowerPoint.GeometricShapeType.rightArrow)"),
        ("insertChevron", "Insert Chevron", "ShapesChevron1", "insertShape(PowerPoint.GeometricShapeType.chevron)"),
        ("insertLineRibbon", "Insert Line", "ShapesLine", "insertLine()"),
        ("insertTextBoxRibbon", "Insert Text Box", "InsertTextBox", "insertTextBox()"),
        ("insertCircle", "Insert Numbered Circle", "ShapesCircle", 'insertNumberedCircle("1")'),
        ("connectShapes", "Connect Two Shapes", "ConnectObjects", "insertConnector()"),
        ("groupRow", "Lay Out as a Row", "GroupAsRows", 'groupAsLayout("row")'),
        ("groupColumn", "Lay Out as a Column", "GroupAsColumns", 'groupAsLayout("column")'),
    ]),
    ("Arrange", "Swap & Apply", [
        ("swapPositionRibbon", "Swap Position", "SwapPosition", "swapPosition()"),
        ("swapSizeRibbon", "Swap Size", "MatchSize", "swapSize()"),
        ("swapFillRibbon", "Swap Fill & Outline", "SwapFillAndOutline", "swapFillAndOutline()"),
        ("pickUpRibbon", "Pick Up Size & Position", "SizePositionCopy", "pickUpSizePosition()"),
        ("applySizeRibbon", "Apply Size & Position", "SizePositionPaste", "applySizePosition()"),
        ("applyMatchingRibbon", "Apply to Matching Objects", "SizePositionMatchingPaste", "applyToMatchingObjects()"),
    ]),
    ("Tables", "Tables", [
        ("addTableRowRibbon", "Add Table Row", "AddRowBottom", "addTableRow()"),
        ("addTableColumnRibbon", "Add Table Column", "AddColumnRight", "addTableColumn()"),
        ("deleteTableRowRibbon", "Delete Last Row", "RemoveLastRow", "deleteLastTableRow()"),
        ("deleteTableColumnRibbon", "Delete Last Column", "RemoveLastColumn", "deleteLastTableColumn()"),
        ("tableToTextRibbon", "Convert Table to Text", "TableToText", "tableToText()"),
    ]),
    ("Special", "Special Shapes", [
        ("specialTitle", "Insert Slide Title", "SlideTitle", 'insertSpecialShape("title")'),
        ("specialConclusion", "Insert Conclusion", "Conclusion", 'insertSpecialShape("conclusion")'),
        ("specialFootnote", "Insert Footnote", "Footnote", 'insertSpecialShape("footnote")'),
        ("specialGhost", "Insert Ghost", "Ghost", 'insertSpecialShape("ghost")'),
        ("specialLabel", "Insert Label", "Labels", 'insertSpecialShape("label")'),
    ]),
    ("Slides", "Slides", [
        ("insertTocRibbon", "Insert Table of Contents", "TableOfContents", "insertTableOfContents()"),
        ("addNoteRibbon", "Add Sticky Note", "AddNote", 'addStickyNote("")'),
    ]),
]

IMPORTS = {
    "../lib/alignment": ["align", "distribute"],
    "../lib/sizePosition": ["matchSize", "stretchToEdge", "straightenLines", "fillGap", "unifyShapes"],
    "../lib/selectSame": ["selectSame"],
    "../lib/visibility": ["setSelectedShapesVisible", "showAllShapes"],
    "../lib/text": ["setAutoSize", "setWordWrap", "setMargins", "setParagraphAlignment",
                    "clearLineBreaks", "clearText", "setBullets", "applyTextStyle",
                    "mergeText", "splitText"],
    "../lib/shapes": ["insertShape", "insertLine", "insertTextBox", "insertNumberedCircle",
                      "insertConnector", "groupAsLayout"],
    "../lib/transform": ["swapPosition", "swapSize", "swapFillAndOutline", "pickUpSizePosition",
                         "applySizePosition", "applyToMatchingObjects"],
    "../lib/tables": ["addTableRow", "addTableColumn", "deleteLastTableRow",
                      "deleteLastTableColumn", "tableToText"],
    "../lib/specialShapes": ["insertSpecialShape"],
    "../lib/slides": ["insertTableOfContents", "addStickyNote"],
}


def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def gen_extension_point():
    """The CustomTab block plus the existing Home-tab button."""
    groups_xml = []
    for gid, glabel, buttons in GROUPS:
        group_icon = buttons[0][2]
        controls = []
        for bid, label, icon, _expr in buttons:
            controls.append(f"""              <Control xsi:type="Button" id="{bid}">
                <Label resid="Lbl.{bid}" />
                <Supertip>
                  <Title resid="Lbl.{bid}" />
                  <Description resid="Lbl.{bid}" />
                </Supertip>
                <Icon>
                  <bt:Image size="16" resid="Ic.{icon}" />
                  <bt:Image size="32" resid="Ic.{icon}" />
                  <bt:Image size="80" resid="Ic.{icon}" />
                </Icon>
                <Action xsi:type="ExecuteFunction">
                  <FunctionName>{bid}</FunctionName>
                </Action>
              </Control>""")
        groups_xml.append(f"""            <Group id="StudioTools.{gid}">
              <Label resid="Grp.{gid}" />
              <Icon>
                <bt:Image size="16" resid="Ic.{group_icon}" />
                <bt:Image size="32" resid="Ic.{group_icon}" />
                <bt:Image size="80" resid="Ic.{group_icon}" />
              </Icon>
{chr(10).join(controls)}
            </Group>""")
    return f"""          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="MacToolsGroup">
                <Label resid="Group.Label" />
                <Icon>
                  <bt:Image size="16" resid="Icon.16" />
                  <bt:Image size="32" resid="Icon.32" />
                  <bt:Image size="80" resid="Icon.80" />
                </Icon>
                <Control xsi:type="Button" id="OpenTaskpaneButton">
                  <Label resid="Taskpane.Label" />
                  <Supertip>
                    <Title resid="Taskpane.Label" />
                    <Description resid="Taskpane.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16" />
                    <bt:Image size="32" resid="Icon.32" />
                    <bt:Image size="80" resid="Icon.80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>MacToolsTaskpane</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
            <CustomTab id="StudioTools.Tab">
              <Group id="StudioTools.Pane">
                <Label resid="Grp.Pane" />
                <Icon>
                  <bt:Image size="16" resid="Icon.16" />
                  <bt:Image size="32" resid="Icon.32" />
                  <bt:Image size="80" resid="Icon.80" />
                </Icon>
                <Control xsi:type="Button" id="OpenTaskpaneFromTab">
                  <Label resid="Taskpane.Label" />
                  <Supertip>
                    <Title resid="Taskpane.Label" />
                    <Description resid="Taskpane.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16" />
                    <bt:Image size="32" resid="Icon.32" />
                    <bt:Image size="80" resid="Icon.80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>MacToolsTaskpane</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url" />
                  </Action>
                </Control>
              </Group>
{chr(10).join(groups_xml)}
              <Label resid="StudioTools.TabLabel" />
            </CustomTab>
          </ExtensionPoint>"""


def gen_resources(host):
    images = ['        <bt:Image id="Icon.16" DefaultValue="{h}/assets/icon-16.png" />',
              '        <bt:Image id="Icon.32" DefaultValue="{h}/assets/icon-32.png" />',
              '        <bt:Image id="Icon.80" DefaultValue="{h}/assets/icon-80.png" />']
    icons = []
    for _gid, _gl, buttons in GROUPS:
        for _bid, _label, icon, _expr in buttons:
            icons.append(icon)
    for icon in sorted(set(icons)):
        images.append(f'        <bt:Image id="Ic.{icon}" DefaultValue="{{h}}/assets/icons/{icon}.png" />')
    images = [i.format(h=host) for i in images]

    short = [
        '        <bt:String id="GetStarted.Title" DefaultValue="StudioTools" />',
        '        <bt:String id="Group.Label" DefaultValue="StudioTools" />',
        '        <bt:String id="Taskpane.Label" DefaultValue="StudioTools" />',
        '        <bt:String id="StudioTools.TabLabel" DefaultValue="Tools" />',
        '        <bt:String id="Grp.Pane" DefaultValue="Task Pane" />',
    ]
    for gid, glabel, buttons in GROUPS:
        short.append(f'        <bt:String id="Grp.{gid}" DefaultValue="{esc(glabel)}" />')
        for bid, label, _icon, _expr in buttons:
            short.append(f'        <bt:String id="Lbl.{bid}" DefaultValue="{esc(label)}" />')

    long = [
        '        <bt:String id="GetStarted.Description" DefaultValue="Open the StudioTools task pane — enhanced features for streamlined presentation workflows." />',
        '        <bt:String id="Taskpane.Tooltip" DefaultValue="Open the StudioTools task pane." />',
    ]
    urls = [
        f'        <bt:Url id="Commands.Url" DefaultValue="{host}/commands.html" />',
        f'        <bt:Url id="Taskpane.Url" DefaultValue="{host}/taskpane.html" />',
        f'        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="{host}/" />',
    ]
    return f"""    <Resources>
      <bt:Images>
{chr(10).join(images)}
      </bt:Images>
      <bt:Urls>
{chr(10).join(urls)}
      </bt:Urls>
      <bt:ShortStrings>
{chr(10).join(short)}
      </bt:ShortStrings>
      <bt:LongStrings>
{chr(10).join(long)}
      </bt:LongStrings>
    </Resources>"""


MANIFEST_TEMPLATE = """<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
  xsi:type="TaskPaneApp">

  <Id>8545941f-5483-4dd7-a665-b13ab3ba43ca</Id>
  <Version>0.1.0.0</Version>
  <ProviderName>Creative Studio</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="StudioTools" />
  <Description DefaultValue="Enhanced features for macOS and web — built for streamlined presentation workflows." />
  <IconUrl DefaultValue="{host}/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="{host}/assets/icon-64.png" />
  <SupportUrl DefaultValue="{host}/" />

  <Hosts>
    <Host Name="Presentation" />
  </Hosts>

  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.5" />
    </Sets>
  </Requirements>

  <DefaultSettings>
    <SourceLocation DefaultValue="{host}/taskpane.html" />
  </DefaultSettings>

  <Permissions>ReadWriteDocument</Permissions>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Presentation">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title" />
            <Description resid="GetStarted.Description" />
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl" />
          </GetStarted>
          <FunctionFile resid="Commands.Url" />
{extension_point}
        </DesktopFormFactor>
      </Host>
    </Hosts>

{resources}
  </VersionOverrides>
</OfficeApp>
"""


def write_manifest(path, host):
    xml = MANIFEST_TEMPLATE.format(
        host=host,
        extension_point=gen_extension_point(),
        resources=gen_resources(host),
    )
    with open(path, "w") as f:
        f.write(xml)
    print("wrote", path)


def gen_commands():
    import_lines = []
    for module, names in IMPORTS.items():
        import_lines.append("import {\n  " + ",\n  ".join(names) + f",\n}} from \"{module}\";")
    runners = []
    for _gid, _gl, buttons in GROUPS:
        for bid, _label, _icon, expr in buttons:
            runners.append(f"  {bid}: () => {expr},")
    return f"""/* Ribbon command handlers — generated by scripts/gen-ribbon.py. Do not edit. */
{chr(10).join(import_lines)}

const RUNNERS: Record<string, () => Promise<void>> = {{
{chr(10).join(runners)}
}};

Office.onReady(() => undefined);

for (const [name, run] of Object.entries(RUNNERS)) {{
  Office.actions.associate(name, (event: Office.AddinCommands.Event) => {{
    run()
      .catch((error) => console.error(`StudioTools: ${{name}} failed`, error))
      .finally(() => event.completed());
  }});
}}
"""


def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    os.chdir(root)
    write_manifest("manifest.xml", "https://jens-lischka.github.io/MacTools")
    write_manifest("manifest.local.xml", "https://localhost:3000")
    with open("src/commands/commands.ts", "w") as f:
        f.write(gen_commands())
    print("wrote src/commands/commands.ts")
    total = sum(len(b) for _g, _l, b in GROUPS)
    print(f"{total} ribbon buttons across {len(GROUPS)} groups")


if __name__ == "__main__":
    main()
