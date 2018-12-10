# NSE-Value-Update

## About
<b>Scrape everyday data from NSE Servers</b> and append them into the <b>excel file.</b><br>
The data can now be worked upon easily by a common user.

## Why ?
The NSE Website allows to view only one day historical data per time.<br>
This poses a problem for the user to study all the values.<br>
My project, appends the values at user specified location, thus able to view the values all together.

## Working
<ol>
  <li>Finds the <b>Dates</b>, whose NSE values are <b>missing</b> from the excel file.
  <li>For each Date (Skipping Public Holidays and Sat-Sun, when market is closed):
    <ul type="disc">
      <li>Scrapes, <b>FIIs Participant Data.</b>
      <li>Scrapes, <b>DIIs Participant Data</b>
      <li>Scrapes, <b>NIFTY Index Data.</b>
      <li>Appends them at right positions.
    </ul>
</ol>

## Working Model
<a href="https://ibb.co/8cz21xR"><img src="https://i.ibb.co/6Bvbjtp/NSE-Index-1.png" alt="NSE-Index-1" border="0" width="60%"></a>

## Next Update Log
<ul>
  <li>Do, the <b>mathematical calculations</b> on above data and append them into the file.
  <li>More <b>User Friendly</b> and <b>faster.</b>
  <li><b>Google Slides</b> Integration.
  <li>More varied Data Addition into the file.
</ul>

## Contributions
If you have any method, to make the running time slower, you can <b>Create a branch</b>, with required Solution for further review.

## Date Project Started
25th November, 2018
  
