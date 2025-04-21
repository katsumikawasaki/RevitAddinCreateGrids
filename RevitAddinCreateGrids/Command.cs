#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;

#endregion

namespace RevitAddinCreateGrids
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            View view = doc.ActiveView;
            
            XYZ basePointMM = new XYZ(0, 0, 0);//基準点
            int xCount = 5;//X軸通芯の数
            int yCount = 3;//Y軸通芯の数
            double xSpanMM = 7200;//Xスパン7200mm
            double ySpanMM = 7200;//Yスパン7200mm

            //通芯を保存するリスト
            List<Grid> xGrids = new List<Grid>();
            List<Grid> yGrids = new List<Grid>();

            using (Transaction tx = new Transaction(doc))
            {
                //通芯を作成するトランザクション
                tx.Start("Transaction Create Grids");
                try
                {
                    //まとめて通芯作成して返す
                    (xGrids,yGrids) = CreateGrids(basePointMM, xCount, yCount, xSpanMM, ySpanMM, doc);
                    tx.Commit();
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("エラー", ex.Message);
                    tx.RollBack();
                    return Result.Failed;
                }
            }
            using (Transaction tx = new Transaction(doc))
            {
                //【重要】通芯はトランザクションが完了してモデルに追加された後でなければ通芯つまりGridのGeometricReferenceを取得できないので、トランザクションを分ける必要がある。
                //寸法線を作成するトランザクション開始
                tx.Start("Transaction Create Dimension");
                try
                {
                    //まとめて寸法作成
                    CreateAllDimensions(doc, xGrids, yGrids, @"点線 - 塗り潰し - 2.5 mm Arial");
                    tx.Commit();
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("エラー", ex.Message);
                    tx.RollBack();
                    return Result.Failed;
                }
            }
            return Result.Succeeded;
        }
        //全ての寸法線を作成する
        public static void CreateAllDimensions(Document doc, List<Grid> xGrids, List<Grid> yGrids,string dimensionTypeName)
        {
            // 寸法線のスタイルを取得
            DimensionType dimensionType = GetDimensionType(doc, dimensionTypeName);

            // 各通芯間の寸法線を作成
            if (xGrids.Count > 1)
            {
                for (int i = 0; i < xGrids.Count - 1; i++)
                {
                    CreateDimension(doc, xGrids[i], xGrids[i + 1], dimensionType, 1200);
                }
            }
            if (yGrids.Count > 1)
            {
                for (int i = 0; i < yGrids.Count - 1; i++)
                {
                    CreateDimension(doc, yGrids[i], yGrids[i + 1], dimensionType, 1200);
                }
            }
            // 端部通芯間の寸法線を作成。通芯が2本以上ある場合のみ作成
            if (xGrids.Count > 2)
            {
                CreateDimension(doc, xGrids[0], xGrids[xGrids.Count - 1], dimensionType, 500);
            }
            if (yGrids.Count > 2)
            {
                CreateDimension(doc, yGrids[0], yGrids[yGrids.Count - 1], dimensionType, 500);
            }
        }
        //通芯を作る
        public static (List<Grid>, List<Grid>) CreateGrids(XYZ basePointMM, int xCount, int yCount, double xSpanMM, double ySpanMM, Document doc)
        {
            //X軸通芯のオフセット
            double xOffset = Conv(-5000);
            //Y軸通芯のオフセット
            double yOffset = Conv(-5000);
            //基準点
            XYZ basePoint = new XYZ(Conv(basePointMM.X), Conv(basePointMM.Y), 0);
            //Xスパン
            double xSpan = Conv(xSpanMM);
            //Yスパン
            double ySpan = Conv(ySpanMM);
            //通芯を保存するリスト
            List<Grid> xGrids = new List<Grid>();
            List<Grid> yGrids = new List<Grid>();

            //X軸通芯作成
            for (int i =0; i < xCount; i++)
            {
                //X軸通芯の開始点
                XYZ start = new XYZ(basePoint.X + xSpan * i, basePoint.Y + yOffset, 0);
                //X軸通芯の終了点
                XYZ end = new XYZ(basePoint.X + xSpan * i, basePoint.Y  + ySpan * yCount, 0);
                //線分作成
                Line line = Line.CreateBound(start, end);
                //通芯作成
                Grid grid = Grid.Create(doc, line);
                //通芯の名前を設定
                grid.Name = "X" + (i + 1).ToString();
                //保存
                xGrids.Add(grid);
            }
            //Y軸通芯作成
            for (int i = 0; i < yCount; i++)
            {
                //Y軸通芯の開始点
                XYZ start = new XYZ(basePoint.X + xOffset, basePoint.Y + ySpan * i, 0);
                //Y軸通芯の終了点
                XYZ end = new XYZ(basePoint.X + xSpan * xCount, basePoint.Y + ySpan * i, 0);
                //線分作成
                Line line = Line.CreateBound(start, end);
                //通芯作成
                Grid grid = Grid.Create(doc, line);
                //通芯の名前を設定
                grid.Name = "Y" + (i + 1).ToString();
                //保存
                yGrids.Add(grid);
            }
            return (xGrids, yGrids);
        }
        //寸法線のスタイルを取得するヘルパーメソッド
        private static DimensionType GetDimensionType(Document doc, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(DimensionType));
            foreach (DimensionType dt in collector)
            {
                if (dt.Name.Equals(typeName))
                {
                    return dt;
                }
            }
            return null;
        }
        // 2つの通芯間に寸法線を作成する関数
        public static void CreateDimension(Document doc, Grid grid1, Grid grid2, DimensionType dimensionType, double offset)
        {
            View view = doc.ActiveView;
            try
            {
                // グリッドのカーブ情報を取得
                Curve curve1 = grid1.Curve;
                Curve curve2 = grid2.Curve;
                //もしカーブ情報がnullならエラー
                if (curve1 == null || curve2 == null)
                {
                    TaskDialog.Show("エラー", "グリッドの曲線情報を取得できませんでした。");
                    return;
                }
                //通芯は直線として扱う
                Line grid1Line = curve1 as Line;
                Line grid2Line = curve2 as Line;
                // ReferenceArrayを使用
                ReferenceArray refArray = new ReferenceArray();
                // グリッド線の参照を取得する
                Reference ref1 = new Reference(grid1);
                Reference ref2 = new Reference(grid2);
                if (ref1 == null || ref2 == null)
                {
                    TaskDialog.Show("エラー", "グリッドの参照を取得できませんでした。");
                    return;
                }
                refArray.Append(ref1);
                refArray.Append(ref2);
                // 寸法線の配置位置を決定
                XYZ grid1Start = grid1Line.GetEndPoint(0);
                XYZ grid1End = grid1Line.GetEndPoint(1);
                XYZ grid2Start = grid2Line.GetEndPoint(0);
                // グリッドの方向ベクトル
                XYZ gridDirection = (grid1End - grid1Start).Normalize();
                // 寸法線の方向ベクトル（通芯と直交する方向）gridDirectionを90度反時計方向に回転
                //マトリクス(cos90, -sin90, sin90, cos90)(x,y)を使って回転させる
                XYZ perpendicular = new XYZ(-gridDirection.Y, gridDirection.X, 0).Normalize();

                // 寸法線の配置位置
                // オフセットoffsetは通芯記号からの距離。offsetプラスは内側、マイナスは外側
                //dimensionLineは位置と方向が重要で長さは関係しない模様。実際の長さはrefArrayで決まる
                Line dimensionLine;
                if (gridDirection.X == 1 || gridDirection.X == -1)
                {
                    // Y軸通芯の場合（X方向に伸びている通芯）
                    dimensionLine = Line.CreateBound(
                        new XYZ(grid1Start.X + Conv(offset), 0, 0), 
                        new XYZ(grid1Start.X + Conv(offset), 1, 0));
                }
                else
                {
                    // X軸通芯の場合
                    dimensionLine = Line.CreateBound(
                        new XYZ(0, grid1Start.Y + Conv(offset), 0), 
                        new XYZ(1, grid1Start.Y + Conv(offset), 0));
                }
                // 寸法線の作成
                Dimension dimension = doc.Create.NewDimension(view, dimensionLine, refArray, dimensionType);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("エラー", $"寸法線の作成中にエラーが発生しました: {ex.Message}");
            }
        }
        //ミリメートル単位を内部単位に変換
        private static double Conv(double mm)
        {
            return UnitUtils.ConvertToInternalUnits(mm, UnitTypeId.Millimeters);
        }
        
    }
}
