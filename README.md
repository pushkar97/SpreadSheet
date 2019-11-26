# SpreadSheet
Apache POI based convenience class for creating excel spreadsheets

It Supports reading and writing data to .xls and .xlsx files

It also supports basic styling with html like syntax.<br/>
Append followning tags to data in each cell for styling.<br/>
&lt;POI-BOLD/&gt;, &lt;POI-ITALIC/&gt;, &lt;POI-STRIKEOUT/&gt;, &lt;POI-UNDERLINE/&gt;<br/>
&lt;POI-BGCOLOR="&lt;color&gt;"/&gt;, &lt;POI-FONTCOLOR="&lt;color&gt;"/&gt;<br/>
it supports poi based indexed colors. for list of colors refer : https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/IndexedColors.html
<br/>
&lt;POI-BORDER [ALL,TOP,LEFT,RIGHT,BOTTOM]="&lt;BorderStyle&gt;" [COLOR,COLOR_TOP,COLOR_LEFT,COLOR_RIGHT,COLOR_BOTTOM]="&lt;IndexedColors&gt;"/&gt;<br/>
For list of supported border styles refer:
https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/BorderStyle.html

It mostly focuses on handling edge cases, such as handling nulls and empty cells, data conversion etc.
