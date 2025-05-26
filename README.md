# KML-CPM-GoogleAppsScript

CPM stands for Copy Paste Machine a KML data composer.

Must be used after interaction with '<a hrev="https://github.com/bostonsinaga/KML-TwinCheck-WebPage">KML-CPM-WebPage</a>'.<br/>
The input data for cell generated from that page in JSON form.

Run 'susun.gs' when JSON data is pasted in the specified cell.

# How to use in Google Spreadsheets

<i>!!! ATTENTION !!!</i><br/>
<i>If you switch to another sheet please reload the Apps Script page.<br/>
Because the target sheet is not detected and causes the target execution to become the previous sheet.<br/>
This applies to all '.gs' codes of this repository.</i>

<h2>Data Constructor for KML Copy Paste Machine</h2>
<ol>
  <li>
    Input in the form of JSON Array code obtained from
    '<a hrev="https://github.com/bostonsinaga/KML-TwinCheck-WebPage">KML-CPM-WebPage</a>'.
  </li>
  <li>
    Input Methods:<br/>
    <ul>
      <li>SINGLE. Paste the JSON in 'A4'.</li>
      <li>PARTS / MULTIPLE (occurs if the number of JSON characters exceeds 50,000 characters).<br/>
        Write the '*' sign in 'A4' to indicate compound input.
        Then paste the JSON parts in 'A5', 'A6', 'An', ...<br/>
        sequentially as many as the number of parts.
      </li>
    </ul>
  </li>
  <li>Make sure all columns in that row are empty because the existing data will be overwritten.</li>
  <li>Then run this script.</li>
</ol>

