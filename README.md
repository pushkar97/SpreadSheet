# SpreadSheet
Apache POI based convenience class for creating excel spreadsheets

It Supports reading and writing data to .xls and .xlsx files

It also supports basic styling with html like syntax.</br>
Append followning tags to data in each cell for styling.</br>
<ul>
  <li><B>&lt;POI-BOLD/&gt;</B> : Adds <b>bold</b> style to cell</li>
  <li><B>&lt;POI-ITALIC/&gt;</B> : Adds <i>italic</i> style to cell</li>
  <li><B>&lt;POI-STRIKEOUT/&gt;</B> : Adds <del>strikeout</del> style to cell</li>
  <li><B>&lt;POI-UNDERLINE/&gt;</B> : Adds <ins>underline</ins> style to cell</li>
  <li><B>&lt;POI-BGCOLOR="&lt;IndexedColor&gt;"/&gt;</B> : Sets background color for cell with specified IndexedColor</li>
  <li><B>&lt;POI-FONTCOLOR="&lt;IndexedColor&gt;"/&gt;</B> : Sets Font color for cell with specified IndexedColor</li>
  <li><b>&lt;POI-BORDER [ALL|TOP,LEFT,RIGHT,BOTTOM]="&lt;BorderStyle&gt;" [COLOR|COLOR_TOP,COLOR_LEFT,COLOR_RIGHT,COLOR_BOTTOM]="&lt;IndexedColors&gt;"/&gt;</b> adds border style to cell</li>
  <li><b>&lt;POI-HYPERLINK TYPE="&lt;HyperlinkType&gt;" URL="&lt;hyperlink&gt;"/&gt;</b></li>
</ul>

for list of IndexedColors refer : https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/IndexedColors.html
</br>
For list of supported border styles refer: https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/BorderStyle.html
</br>
For list of supported Hyperlink types refer : https://poi.apache.org/apidocs/dev/org/apache/poi/common/usermodel/HyperlinkType.html




It mostly focuses on handling edge cases, such as handling nulls and empty cells, data conversion etc.
