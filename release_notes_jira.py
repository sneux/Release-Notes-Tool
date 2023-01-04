import requests
import json
from docx import Document
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Inches, RGBColor
from docx.oxml.ns import qn
from docx.oxml.shared import OxmlElement, qn


def get_lists(version: str) -> str:
  version = version
  url = f"http://dataremote.atlassian.net/rest/api/2/search?jql=project in (CXX,VOIP,DEV) AND fixVersion in (\"{version}\" ) ORDER BY project, updated DESC"
  payload = ""
  headers = {
  'Authorization': 'Basic #'
}
  response = requests.request("GET", url, headers=headers, data=payload)
  data = json.loads(response.text)          # all of the data inside the JSON file
  issues = data['issues']                   # a list of dictionaries containing the info we need
  bug_list = []
  other_list = []

  # we have to iterate over the list of dictionaries and manually find the info we need
  index = 0
  for item in issues:
    key = (issues[index]['key'])
    type = (issues[index]['fields']['issuetype']['name'])
    
    # because we are searching for multiple custom fields across different projects - 
    # we have to first see if the custom fields exist for the project and if not
    # we pass the errror it raises to keep iterating 
    if type == 'Bug':
      try:
        release_n = issues[index]['fields']['customfield_10691']
      except (KeyError):
        try:
          release_n = issues[index]['fields']['customfield_10615']
        except (KeyError):
          pass 
      bug_list.append((key,release_n))
    
    else:
      try:
        release_n = issues[index]['fields']['customfield_10691']
      except (KeyError):
        try:
          release_n = issues[index]['fields']['customfield_10615']
        except (KeyError):
          release_n = "N/A"  
      other_list.append((key,release_n))
    
    index += 1
  return other_list, bug_list

# we need to get rid of all the "N/A" entries
def clean_list(a_list):
  test_string = "N/A"
  new_list = [item for item in a_list if test_string not in item]
  return new_list

# def fill_bg(cell):

def make_table(doc, header, a_list):
  table = doc.add_table(rows = 1, cols = 1, style = "Table Grid")
  table_header = table.rows[0].cells

  row = table.rows[0].cells
  header_cell = row[0]
  table_header[0].text = header
  h_bg = header_cell._tc.get_or_add_tcPr()
  hAlign = OxmlElement("w:shd")
  hAlign.set(qn("w:fill"), "#C00000")
  h_bg.append(hAlign)
  
  table = doc.add_table(rows = 1, cols = 2, style = "Table Grid")
  table.allow_autofit = False
  table.rows[0].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

  row = table.rows[0].cells
  cell, cell2 = row[0], row[1]

  row[0].text = "ID#"
  cell_bg = cell._tc.get_or_add_tcPr()
  tcVAlign = OxmlElement("w:shd")
  tcVAlign.set(qn("w:fill"), "#262626")

  row[1].text = "Description"
  cell_bg2 = cell2._tc.get_or_add_tcPr()
  tcVAlign2 = OxmlElement("w:shd")
  tcVAlign2.set(qn("w:fill"), "#262626")

  cell_bg.append(tcVAlign)
  cell_bg2.append(tcVAlign2)

  for id in a_list:
    row = table.add_row().cells
    row[0].text = id[0]
    row[1].text = id[1]
  for cell in table.columns[1].cells:
      cell.width = Inches(10)

  doc.add_page_break()


def make_doc(other_list, bug_list):
  #first let's create and format the first page 
  doc = Document()
  doc.add_picture('dataremotelogo.png')
  last_paragraph = doc.paragraphs[-1]
  last_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

  style = doc.styles['Normal']
  font = style.font
  font.name = "Times New Roman"
  p = doc.add_paragraph()
  header = doc.add_heading('', 0)
  header.add_run("VAB-1").font.color.rgb = RGBColor(0,0,0)
  header.alignment = 1
  header = doc.add_heading('', 0)
  header.add_run("<FIRMARE Version>").font.color.rgb = RGBColor(0,0,0)
  header.alignment = 1
  header = doc.add_heading('', 0)
  header.add_run("Release Notes").font.color.rgb = RGBColor(0,0,0)
  header.alignment = 1
  rFonts = header.style.element.rPr.rFonts
  rFonts.set(qn("w:asciiTheme"), "Times New Roman")
  doc.add_page_break()

  # now lets add the header image
  header = doc.sections[0].header
  paragraph = header.paragraphs[0]
  logo_header = paragraph.add_run()
  logo_header.add_picture("dataremotelogo.png", width=Inches(3))

  # now we add the software new features section
  new_features_header = doc.add_heading('')
  new_features_header.add_run("Release Notes").font.color.rgb = RGBColor(0,0,0)
  rFonts = new_features_header.style.element.rPr.rFonts
  rFonts.set(qn("w:asciiTheme"), "Times New Roman")
  nf_desc = doc.add_paragraph('This section identifies new software features on this release.')
  # creating the table
  table = make_table(doc, "New Features", other_list)

  # now the page with "known issues" table
  resolved_issues_header = doc.add_heading('')
  resolved_issues_header.add_run("Software Resolved Issues").font.color.rgb = RGBColor(0,0,0)
  rFonts = resolved_issues_header.style.element.rPr.rFonts
  rFonts.set(qn("w:asciiTheme"), "Times New Roman")
  ri_desc = doc.add_paragraph('This section identifies issues that have been resolved since the last release of the software.')
  table2 = make_table(doc, "Resolved Issues", bug_list)

  # make it so the header image doesn't show on the first page
  for section in doc.sections:
    section.different_first_page_header_footer = True

  doc.save('Release Notes.docx')

def main():
  bug_list, other_list = get_lists("ATT_R3.1")
  bug_list = clean_list(bug_list)
  other_list = clean_list(other_list)
  make_doc(other_list, bug_list)

if __name__ == "__main__":
    main()


