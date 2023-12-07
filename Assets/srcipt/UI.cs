using System.Collections;
using System.Collections.Generic;
using TMPro;
using Unity.VisualScripting;
using UnityEngine;
using OfficeOpenXml;
using System.IO;
using System.Linq;
using UnityEngine.UI;
using System.Drawing;
using System;
using Random = UnityEngine.Random;
using Color = UnityEngine.Color;

public class UI : MonoBehaviour
{
    // Start is called before the first frame update
    void Start()
    {
        Debug.Log("UI is on");
        GameObject guiti = GameObject.Find("7");
        Color32 newColor = new Color32(255,255,255, 255);
        updateSaveNum();
        alreadySaveUi();
        changeTitle();
    }


    static Color newColor1;
    static Color newColor2;
    static float h1,h2,s1,s2,v1,v2;
    static byte score;
    static int roll = 2;
    static byte sameCont1;
    static byte cabinetFlag =1;
    static int SaveNum1 = 0;
    static int SaveNum2 = 0;
    static int SaveNum3 = 0;


    public void clickButton1()
    {
        randomColor();
        score = 1;
        saveColor();
    }
    public void clickButton2()
    {
        randomColor();
        score = 2;
        saveColor();
    }
    public void clickButton3()
    {
        randomColor();
        score = 3;
        saveColor();
    }
    public void clickButton4()
    {
        randomColor();
        score = 4;
        saveColor();
    }
    public void clickButton5()
    {
        randomColor();
        score = 5;
        saveColor();
    }

    void randomColor()
    {
        h1 = Random.Range(0f, 1f);
        h2 = Random.Range(0f, 1f);
        s1 = Random.Range(0f, 1f);
        s2 = Random.Range(0f, 1f);
        v1 = Random.Range(0f, 1f);
        v2 = Random.Range(0f, 1f);

        newColor1 = Color.HSVToRGB(h1,s1, v1);
        newColor2 = Color.HSVToRGB(h2, s2, v2);

        GameObject door1 = GameObject.Find("1");
        GameObject door2 = GameObject.Find("2");
        GameObject door3 = GameObject.Find("3");
        GameObject cabinet1 = GameObject.Find("4");

        GameObject door1_2 = GameObject.Find("second1");
        GameObject door2_2 = GameObject.Find("second2");
        GameObject door3_2 = GameObject.Find("second3");
        GameObject door4_2 = GameObject.Find("second4");
        GameObject door5_2 = GameObject.Find("second5");
        GameObject door6_2 = GameObject.Find("second6");
        GameObject cabinet2 = GameObject.Find("second7");

        GameObject door1_3 = GameObject.Find("third1");
        GameObject door2_3 = GameObject.Find("third2");
        GameObject door3_3 = GameObject.Find("third3");
        GameObject door4_3 = GameObject.Find("third4");
        GameObject door5_3 = GameObject.Find("third5");
        GameObject cabinet3 = GameObject.Find("third6");

        switch (cabinetFlag)
        {
            case 1:
                cabinet1.GetComponentInChildren<MeshRenderer>().material.color = new Color32(200, 200, 200, 255);
                sameCont1 = (byte)Random.Range(0, 2);
                if (sameCont1 == 0)
                {
                    door1.GetComponentInChildren<MeshRenderer>().material.color = newColor1;
                    door2.GetComponentInChildren<MeshRenderer>().material.color = newColor2;
                    door3.GetComponentInChildren<MeshRenderer>().material.color = newColor2;

                }
                else if (sameCont1 == 1)
                {
                    door1.GetComponentInChildren<MeshRenderer>().material.color = newColor1;
                    door2.GetComponentInChildren<MeshRenderer>().material.color = newColor2;
                    door3.GetComponentInChildren<MeshRenderer>().material.color = newColor1;
                }
                else if (sameCont1 == 2)
                {
                    door1.GetComponentInChildren<MeshRenderer>().material.color = newColor1;
                    door2.GetComponentInChildren<MeshRenderer>().material.color = newColor1;
                    door3.GetComponentInChildren<MeshRenderer>().material.color = newColor2;

                }

                break;
            case 2:
                door1_2.GetComponentInChildren<MeshRenderer>().material.color = newColor1;
                door2_2.GetComponentInChildren<MeshRenderer>().material.color = newColor1;
                door3_2.GetComponentInChildren<MeshRenderer>().material.color = newColor1;
                door4_2.GetComponentInChildren<MeshRenderer>().material.color = newColor1;
                door5_2.GetComponentInChildren<MeshRenderer>().material.color = newColor1;
                door6_2.GetComponentInChildren<MeshRenderer>().material.color = newColor1;
                cabinet2.GetComponentInChildren<MeshRenderer>().material.color = newColor2;
                break;
            case 3:
                door1_3.GetComponentInChildren<MeshRenderer>().material.color = newColor1;
                door2_3.GetComponentInChildren<MeshRenderer>().material.color = newColor1;
                door3_3.GetComponentInChildren<MeshRenderer>().material.color = newColor1;
                door4_3.GetComponentInChildren<MeshRenderer>().material.color = newColor1;
                door5_3.GetComponentInChildren<MeshRenderer>().material.color = newColor1;
                cabinet3.GetComponentInChildren<MeshRenderer>().material.color = newColor2;

                break;

        }



    }

    public static (double Hue, double Saturation, double Value) RgbToHsv(byte red, byte green, byte blue)
    {
        double r = red / 255.0;
        double g = green / 255.0;
        double b = blue / 255.0;

        double max = Math.Max(r, Math.Max(g, b));
        double min = Math.Min(r, Math.Min(g, b));

        double hue, saturation, value;

        if (max == min)
        {
            hue = 0;
        }
        else if (max == r)
        {
            hue = 60 * (0 + (g - b) / (max - min));
        }
        else if (max == g)
        {
            hue = 60 * (2 + (b - r) / (max - min));
        }
        else
        {
            hue = 60 * (4 + (r - g) / (max - min));
        }

        if (hue < 0)
        {
            hue += 360;
        }

        if (max == 0)
        {
            saturation = 0;
        }
        else
        {
            saturation = (max - min) / max;
        }

        value = max;

        return (hue, saturation, value);
    }





    void updateSaveNum()
    {
        string filename = "E:/data1.xlsx";
        if (cabinetFlag == 1)
        {
            filename = "E:/data1.xlsx";
            FileInfo fileInfo = new FileInfo(filename);

            using (ExcelPackage excel = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = excel.Workbook.Worksheets[1];
                SaveNum1 = int.Parse(worksheet.Cells[2, 11].Value.ToString()) - 2;
            }
        }
        else if (cabinetFlag == 2)
        {
            filename = "E:/data2.xlsx";
            FileInfo fileInfo = new FileInfo(filename);

            using (ExcelPackage excel = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = excel.Workbook.Worksheets[1];
                SaveNum2 = int.Parse(worksheet.Cells[2, 11].Value.ToString()) - 2;
            }
        }
        else if (cabinetFlag == 3)
        {
            filename = "E:/data3.xlsx";
            FileInfo fileInfo = new FileInfo(filename);

            using (ExcelPackage excel = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = excel.Workbook.Worksheets[1];
                SaveNum3 = int.Parse(worksheet.Cells[2, 11].Value.ToString()) - 2;
            }
        }

    }
    void saveColor()
    {        
        string filename = "E:/data1.xlsx";
        if (cabinetFlag ==1)
        {
            filename = "E:/data1.xlsx";
        }
        else if (cabinetFlag ==2)
        {
            filename = "E:/data2.xlsx";
        }
        else if (cabinetFlag == 3)
        {
            filename = "E:/data3.xlsx";
        }
        FileInfo fileInfo = new FileInfo(filename);


        using (ExcelPackage excel = new ExcelPackage(fileInfo))
        {
            ExcelWorksheet worksheet = excel.Workbook.Worksheets[1];
            string s = worksheet.Cells[1,1].Value.ToString();
            Debug.Log(s);

            roll = int.Parse(worksheet.Cells[2, 11].Value.ToString() );            
            worksheet.Cells[roll, 1].Value = h1;
            worksheet.Cells[roll, 2].Value = s1;
            worksheet.Cells[roll, 3].Value = v1;

            worksheet.Cells[roll, 4].Value = h2;
            worksheet.Cells[roll, 5].Value = s2;
            worksheet.Cells[roll, 6].Value = v2;

            worksheet.Cells[roll, 7].Value = score;
            
            roll++;
            worksheet.Cells[2, 11].Value = roll;
            excel.Save();
        }
        updateSaveNum();

        alreadySaveUi();

    }



    public void changeCabinetFlagTo1()
    {
        GameObject camera = GameObject.Find("MainCamera");
        cabinetFlag = 1;
        camera.transform.position = new Vector3(142, 15, 173);
        camera.transform.rotation = Quaternion.Euler(0f,180f,0f);
        updateSaveNum();

        alreadySaveUi();
        changeTitle();
    }
    public void changeCabinetFlagTo2()
    {
        GameObject camera = GameObject.Find("MainCamera");
        cabinetFlag = 2;
        camera.transform.position = new Vector3(-105, 15, 173);
        camera.transform.rotation = Quaternion.Euler(0f, 180f, 0f);
        updateSaveNum();

        alreadySaveUi();
        changeTitle();

    }
    public void changeCabinetFlagTo3()
    {
        GameObject camera = GameObject.Find("MainCamera");
        cabinetFlag = 3;
        camera.transform.position = new Vector3(-326, 15, 173);
        camera.transform.rotation = Quaternion.Euler(0f, 180f, 0f);
        updateSaveNum();

        alreadySaveUi();
        changeTitle();

    }

    public void right()
    {
        GameObject camera = GameObject.Find("MainCamera");
        if (cabinetFlag == 1)
        {
            camera.transform.position = new Vector3(81.05f, 15f, 168.87f);
            camera.transform.rotation = Quaternion.Euler(0f,-200.806f,0f);
        }else if (cabinetFlag == 2)
        {
            camera.transform.position = new Vector3(-191.8f, 15f, 180.4f);
            camera.transform.rotation = Quaternion.Euler(0f, -206.834f, 0f);
        }
        else if (cabinetFlag == 3) {
            camera.transform.position = new Vector3(-395.7f, 15f, 165.2f);
            camera.transform.rotation = Quaternion.Euler(0f, -203.959f, 0f);
        }

    }
    public void left()
    {
        GameObject camera = GameObject.Find("MainCamera");
        if (cabinetFlag == 1)
        {
            camera.transform.position = new Vector3(199.2f, 15f, 165.6f);
            camera.transform.rotation = Quaternion.Euler(0f, -160.027f, 0f);
        }
        else if (cabinetFlag == 2)
        {
            camera.transform.position = new Vector3(-56.4f, 15f, 170.6f);
            camera.transform.rotation = Quaternion.Euler(0f, -162.221f, 0f);
        }
        else if (cabinetFlag == 3)
        {
            camera.transform.position = new Vector3(-278.1f, 15f, 173f);
            camera.transform.rotation = Quaternion.Euler(0f, -163.711f, 0f);
        }

    }

    public void zheng ()
    {
        GameObject camera = GameObject.Find("MainCamera");
        if (cabinetFlag == 1)
        {
            camera.transform.position = new Vector3(142, 15, 173);
            camera.transform.rotation = Quaternion.Euler(0f, 180f, 0f);
        }
        else if (cabinetFlag == 2)
        {
            camera.transform.position = new Vector3(-105, 15, 173);
            camera.transform.rotation = Quaternion.Euler(0f, 180f, 0f);
        }
        else if (cabinetFlag == 3)
        {
            camera.transform.position = new Vector3(-326, 15, 173);
            camera.transform.rotation = Quaternion.Euler(0f, 180f, 0f);
        }
    }
    public void alreadySaveUi()
    {
        GameObject UI = GameObject.Find("tongzhilan");

        if (cabinetFlag == 1)
        {
            UI.GetComponent<TMP_Text>().text = "already save"+SaveNum1;
        }
        else if (cabinetFlag == 2)
        {
            UI.GetComponent<TMP_Text>().text = "already save" + SaveNum2  ;

        }
        else if (cabinetFlag == 3)
        {
            UI.GetComponent<TMP_Text>().text = "already save" + SaveNum3;

        }

    }

    public void changeTitle()
    {
        GameObject title = GameObject.Find ("title");
        if (cabinetFlag == 1)
        {
            title.GetComponent<TMP_Text>().text = "cabinet first";
        }else if (cabinetFlag == 2)
        {
            title.GetComponent<TMP_Text>().text = "cabinet second";

        }
        else
        {
            title.GetComponent<TMP_Text>().text = "cabinet third";

        }
    }
}
