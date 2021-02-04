package com.anandh.readexcelfiledemo

import android.content.res.AssetManager
import android.os.Bundle
import android.util.Log
import android.widget.TextView
import androidx.appcompat.app.AppCompatActivity
import org.apache.poi.hssf.usermodel.HSSFCell
import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.poifs.filesystem.POIFSFileSystem
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.json.JSONException
import org.json.JSONObject
import java.io.InputStream
import java.util.*


class MainActivity : AppCompatActivity() {

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)
    }

    override fun onResume() {
        super.onResume()
        readExcelFileFromAssets()

        val  displayText = findViewById<TextView>(R.id.tvText)
        findViewById<TextView>(R.id.tvDefaultSelected).text = "Selected Language: "+Locale.getDefault().displayLanguage
        displayText.append("\n")
        for (locale in Locale.getAvailableLocales()) {
            displayText.append(locale.displayLanguage+" - ("+locale.toLanguageTag()+") - "+locale.displayName +"\n")
        }
    }

    private fun readExcelFileFromAssets(): String? {
        var returnValues = ""
        var text = findViewById<TextView>(R.id.tvText)
        try {
            var myInput: InputStream
            // initialize asset manager
            val assetManager: AssetManager = assets
            //  open excel sheet
            myInput = assetManager.open("eko_multi_ling_saparate_file.xls");
            // Create a POI File System object
            val myFileSystem = POIFSFileSystem(myInput)
            // Create a workbook using the File System
            val wb = HSSFWorkbook(myFileSystem)
            // Get the first sheet from workbook
            val mySheet = wb.getSheetAt(0)
            // We now need something to iterate through the cells.
            val rowIter: Iterator<Row> = mySheet.rowIterator()
            var rowno = 0
            text.append("\n")
            while (rowIter.hasNext()) {
                val myRow = rowIter.next() as HSSFRow
                if (rowno != 0) {
                    val cellIter: Iterator<Cell> = myRow.cellIterator()
                    var colno = 0
                    var newStringName = ""
                    var designComponentName = ""
                    var actualName = ""
                    while (cellIter.hasNext()) {
                        val myCell = cellIter.next() as HSSFCell
                        when (colno) {
//                            3 -> {
//                                designComponentName = myCell.toString()
//                            }
                            0 -> {
                                newStringName = myCell.toString()
                            }
                            1 -> {
                                actualName = myCell.toString()
                            }
                        }
                        colno++
                    }
//                    val completeString: String =
//                        newStringName + "_" + designComponentWithReplaceHyphenAndSpaces(
//                            designComponentName
//                        )
                    Log.i(
                        "AK-STRINGS:",
                        "<string name=" + "??$" + newStringName + "??$" + "" + ">" + actualName + "</string>"
                    )
//
//                    newStringName = ""
//                    designComponentName = ""
//                    actualName = ""
                }
                rowno++
            }
        } catch (e: Exception) {
            //Log.e("AK-Exception::", "AK $e")
        }
        return returnValues
    }

   private fun readExcelFileForServerError(): String? {
        var returnValues = ""
        var text = findViewById<TextView>(R.id.tvText)
        try {
            var myInput: InputStream
            val assetManager: AssetManager = assets
            myInput = assetManager.open("spanish.xls");
            val myFileSystem = POIFSFileSystem(myInput)
            val wb = HSSFWorkbook(myFileSystem)
            val mySheet = wb.getSheetAt(0)
            val rowIter: Iterator<Row> = mySheet.rowIterator()
            var rowno = 0
            text.append("\n")
            while (rowIter.hasNext()) {
                val myRow = rowIter.next() as HSSFRow
                if (rowno != 0) {
                    val cellIter: Iterator<Cell> = myRow.cellIterator()
                    var colno = 0
                    var newStringName = ""
                    var designComponentName = ""
                    var actualName = ""
                    while (cellIter.hasNext()) {
                        val myCell = cellIter.next() as HSSFCell
                        when (colno) {
                            0 -> {
                                newStringName = myCell.toString()
                            }
                            2 -> {
                                actualName = myCell.toString()
                            }
                        }
                        colno++
                    }
                    Log.i(
                        "AK-STRINGS:",
                        "??$" + newStringName + "??$" + ":" + "??$" + actualName + "??$$$$"
                    )
//                    newStringName = ""
//                    designComponentName = ""
//                    actualName = ""
                }
                rowno++
            }
        } catch (e: Exception) {
            Log.e("AK-STRINGS-Exception::", "AK $e")
        }
        return returnValues
    }

    private fun designComponentWithReplaceHyphenAndSpaces(designComponentString: String): String? {
        return designComponentString.toLowerCase(Locale.getDefault()).replace("-", "_")
            .replace(" ", "_")
    }


    private fun readJSONFromAsset(): String? {
        var json: String? = null
        try {
            val inputStream: InputStream = assets.open("multi_ling.json")
            json = inputStream.bufferedReader().use { it.readText() }

        } catch (ex: Exception) {
            ex.printStackTrace()
            return null
        }
        return json
    }

    private fun readTheFileAndPrintOnConsole(){
        try {
            val obj = JSONObject(readJSONFromAsset()!!)
            // get JSONObject from JSON file
            val language: JSONObject = obj.getJSONObject("German")
            // get default data
            val defaultLanguage =Locale.getDefault().displayLanguage
            Log.i(
                "AK-STRINGS:",language.getString("A valid :type must be passed")
            )
            findViewById<TextView>(R.id.tvText).text = language.getString("A valid :type must be passed")
        } catch (e: JSONException) {
            e.printStackTrace()
        }
    }
}