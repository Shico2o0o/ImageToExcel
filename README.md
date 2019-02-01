# ImageToExcel
A simple Java app inspired by https://makeanddo4d.com/spreadsheet/ by Andrew Taylor.  I just cloned his app into Java language using the POI apache library. The [stand-up comedy](https://www.youtube.com/watch?v=UBX2QQHlQ_I) was made by Matt Parker. The app takes an image as input, and creates an excel file. Open the created excel file and zoom out!

The app makes a spreadsheet with the RGB values of each pixel of the image, and creates a conditional formatting for each row. You can see the image in the spreadsheet when zooming out.


Example:

    Original:
<p>
    <img src="https://github.com/Shico2o0o/ImageToExcel/blob/master/example/Penguins.jpg" width="220" height="240" />
</p>

    Excelled:
<p>
    <img src="https://github.com/Shico2o0o/ImageToExcel/blob/master/example/ExcelledPenguins.PNG" width="220" height="240" />
</p>
