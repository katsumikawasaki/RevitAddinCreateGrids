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
            
            XYZ basePointMM = new XYZ(0, 0, 0);//��_
            int xCount = 5;//X���ʐc�̐�
            int yCount = 3;//Y���ʐc�̐�
            double xSpanMM = 7200;//X�X�p��7200mm
            double ySpanMM = 7200;//Y�X�p��7200mm

            //�ʐc��ۑ����郊�X�g
            List<Grid> xGrids = new List<Grid>();
            List<Grid> yGrids = new List<Grid>();

            using (Transaction tx = new Transaction(doc))
            {
                //�ʐc���쐬����g�����U�N�V����
                tx.Start("Transaction Create Grids");
                try
                {
                    //�܂Ƃ߂Ēʐc�쐬���ĕԂ�
                    (xGrids,yGrids) = CreateGrids(basePointMM, xCount, yCount, xSpanMM, ySpanMM, doc);
                    tx.Commit();
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("�G���[", ex.Message);
                    tx.RollBack();
                    return Result.Failed;
                }
            }
            using (Transaction tx = new Transaction(doc))
            {
                //�y�d�v�z�ʐc�̓g�����U�N�V�������������ă��f���ɒǉ����ꂽ��łȂ���Βʐc�܂�Grid��GeometricReference���擾�ł��Ȃ��̂ŁA�g�����U�N�V�����𕪂���K�v������B
                //���@�����쐬����g�����U�N�V�����J�n
                tx.Start("Transaction Create Dimension");
                try
                {
                    //�܂Ƃ߂Đ��@�쐬
                    CreateAllDimensions(doc, xGrids, yGrids, @"�_�� - �h��ׂ� - 2.5 mm Arial");
                    tx.Commit();
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("�G���[", ex.Message);
                    tx.RollBack();
                    return Result.Failed;
                }
            }
            return Result.Succeeded;
        }
        //�S�Ă̐��@�����쐬����
        public static void CreateAllDimensions(Document doc, List<Grid> xGrids, List<Grid> yGrids,string dimensionTypeName)
        {
            // ���@���̃X�^�C�����擾
            DimensionType dimensionType = GetDimensionType(doc, dimensionTypeName);

            // �e�ʐc�Ԃ̐��@�����쐬
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
            // �[���ʐc�Ԃ̐��@�����쐬�B�ʐc��2�{�ȏ゠��ꍇ�̂ݍ쐬
            if (xGrids.Count > 2)
            {
                CreateDimension(doc, xGrids[0], xGrids[xGrids.Count - 1], dimensionType, 500);
            }
            if (yGrids.Count > 2)
            {
                CreateDimension(doc, yGrids[0], yGrids[yGrids.Count - 1], dimensionType, 500);
            }
        }
        //�ʐc�����
        public static (List<Grid>, List<Grid>) CreateGrids(XYZ basePointMM, int xCount, int yCount, double xSpanMM, double ySpanMM, Document doc)
        {
            //X���ʐc�̃I�t�Z�b�g
            double xOffset = Conv(-5000);
            //Y���ʐc�̃I�t�Z�b�g
            double yOffset = Conv(-5000);
            //��_
            XYZ basePoint = new XYZ(Conv(basePointMM.X), Conv(basePointMM.Y), 0);
            //X�X�p��
            double xSpan = Conv(xSpanMM);
            //Y�X�p��
            double ySpan = Conv(ySpanMM);
            //�ʐc��ۑ����郊�X�g
            List<Grid> xGrids = new List<Grid>();
            List<Grid> yGrids = new List<Grid>();

            //X���ʐc�쐬
            for (int i =0; i < xCount; i++)
            {
                //X���ʐc�̊J�n�_
                XYZ start = new XYZ(basePoint.X + xSpan * i, basePoint.Y + yOffset, 0);
                //X���ʐc�̏I���_
                XYZ end = new XYZ(basePoint.X + xSpan * i, basePoint.Y  + ySpan * yCount, 0);
                //�����쐬
                Line line = Line.CreateBound(start, end);
                //�ʐc�쐬
                Grid grid = Grid.Create(doc, line);
                //�ʐc�̖��O��ݒ�
                grid.Name = "X" + (i + 1).ToString();
                //�ۑ�
                xGrids.Add(grid);
            }
            //Y���ʐc�쐬
            for (int i = 0; i < yCount; i++)
            {
                //Y���ʐc�̊J�n�_
                XYZ start = new XYZ(basePoint.X + xOffset, basePoint.Y + ySpan * i, 0);
                //Y���ʐc�̏I���_
                XYZ end = new XYZ(basePoint.X + xSpan * xCount, basePoint.Y + ySpan * i, 0);
                //�����쐬
                Line line = Line.CreateBound(start, end);
                //�ʐc�쐬
                Grid grid = Grid.Create(doc, line);
                //�ʐc�̖��O��ݒ�
                grid.Name = "Y" + (i + 1).ToString();
                //�ۑ�
                yGrids.Add(grid);
            }
            return (xGrids, yGrids);
        }
        //���@���̃X�^�C�����擾����w���p�[���\�b�h
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
        // 2�̒ʐc�Ԃɐ��@�����쐬����֐�
        public static void CreateDimension(Document doc, Grid grid1, Grid grid2, DimensionType dimensionType, double offset)
        {
            View view = doc.ActiveView;
            try
            {
                // �O���b�h�̃J�[�u�����擾
                Curve curve1 = grid1.Curve;
                Curve curve2 = grid2.Curve;
                //�����J�[�u���null�Ȃ�G���[
                if (curve1 == null || curve2 == null)
                {
                    TaskDialog.Show("�G���[", "�O���b�h�̋Ȑ������擾�ł��܂���ł����B");
                    return;
                }
                //�ʐc�͒����Ƃ��Ĉ���
                Line grid1Line = curve1 as Line;
                Line grid2Line = curve2 as Line;
                // ReferenceArray���g�p
                ReferenceArray refArray = new ReferenceArray();
                // �O���b�h���̎Q�Ƃ��擾����
                Reference ref1 = new Reference(grid1);
                Reference ref2 = new Reference(grid2);
                if (ref1 == null || ref2 == null)
                {
                    TaskDialog.Show("�G���[", "�O���b�h�̎Q�Ƃ��擾�ł��܂���ł����B");
                    return;
                }
                refArray.Append(ref1);
                refArray.Append(ref2);
                // ���@���̔z�u�ʒu������
                XYZ grid1Start = grid1Line.GetEndPoint(0);
                XYZ grid1End = grid1Line.GetEndPoint(1);
                XYZ grid2Start = grid2Line.GetEndPoint(0);
                // �O���b�h�̕����x�N�g��
                XYZ gridDirection = (grid1End - grid1Start).Normalize();
                // ���@���̕����x�N�g���i�ʐc�ƒ�����������jgridDirection��90�x�����v�����ɉ�]
                //�}�g���N�X(cos90, -sin90, sin90, cos90)(x,y)���g���ĉ�]������
                XYZ perpendicular = new XYZ(-gridDirection.Y, gridDirection.X, 0).Normalize();

                // ���@���̔z�u�ʒu
                // �I�t�Z�b�goffset�͒ʐc�L������̋����Boffset�v���X�͓����A�}�C�i�X�͊O��
                //dimensionLine�͈ʒu�ƕ������d�v�Œ����͊֌W���Ȃ��͗l�B���ۂ̒�����refArray�Ō��܂�
                Line dimensionLine;
                if (gridDirection.X == 1 || gridDirection.X == -1)
                {
                    // Y���ʐc�̏ꍇ�iX�����ɐL�тĂ���ʐc�j
                    dimensionLine = Line.CreateBound(
                        new XYZ(grid1Start.X + Conv(offset), 0, 0), 
                        new XYZ(grid1Start.X + Conv(offset), 1, 0));
                }
                else
                {
                    // X���ʐc�̏ꍇ
                    dimensionLine = Line.CreateBound(
                        new XYZ(0, grid1Start.Y + Conv(offset), 0), 
                        new XYZ(1, grid1Start.Y + Conv(offset), 0));
                }
                // ���@���̍쐬
                Dimension dimension = doc.Create.NewDimension(view, dimensionLine, refArray, dimensionType);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("�G���[", $"���@���̍쐬���ɃG���[���������܂���: {ex.Message}");
            }
        }
        //�~�����[�g���P�ʂ�����P�ʂɕϊ�
        private static double Conv(double mm)
        {
            return UnitUtils.ConvertToInternalUnits(mm, UnitTypeId.Millimeters);
        }
        
    }
}
