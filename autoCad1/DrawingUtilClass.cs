using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using System.Windows.Forms;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.Geometry;
using System.Collections.Generic;
using System;
using System.Threading.Tasks;
using System.Data;

using System.Data.SQLite;

namespace autoCad1
{
    public class DrawingUtilClass
    {

        static Document document = Application.DocumentManager.MdiActiveDocument;
        Database dataBase = document.Database;
        Editor editor = document.Editor;

        [CommandMethod("GetInfo")]
        public void GetInfo()
        {
            ChooseElements();
        }

        public void ChooseElements()
        {
            int countElements = 0;
            Dictionary<string, int> dictionaryOfElementNumbers = new Dictionary<string, int>();
            Dictionary<string, int> dictionaryOfElementsQuantity = new Dictionary<string, int>();
            //удаление предыдущих доп. блоков с информацией на чертеже
            EraseBlocks(dataBase);

            using (Transaction transaction = dataBase.TransactionManager.StartTransaction())
            {
                editor.WriteMessage("Выберете элементы блока!"); 
                PromptSelectionResult selectionResult = editor.GetSelection();//выбор объектов

                if (selectionResult.Status == PromptStatus.OK)
                {
                    SelectionSet selectionSet = selectionResult.Value;
                   //перебор выбранных объектов
                    foreach (SelectedObject selectedObject in selectionSet)
                    {
                        if (selectedObject != null)
                        {
                            // открываем выбранный объект в виде сущности
                            Entity objEntity = transaction.GetObject(selectedObject.ObjectId, OpenMode.ForRead) as Entity;
                           
                            if (objEntity != null)
                            {
                                ResultBuffer xData = objEntity.XData; //получаем данные свойства xData для нахождения GUID
                                if (xData != null)
                                {
                                    // извлекаем GUID из XData
                                    string guidString = GetGuidFromXData(xData);
                                    if (!string.IsNullOrEmpty(guidString))
                                    {
                                        //добавление информации о кол-ве элементов и номеров порядка элементов в соответствующие словари
                                        if (dictionaryOfElementNumbers.ContainsKey(guidString))
                                        {
                                            int prevQuantityValue = dictionaryOfElementsQuantity[guidString];
                                            prevQuantityValue += 1;
                                            dictionaryOfElementsQuantity[guidString] = prevQuantityValue; 
                                        }
                                        else
                                        {
                                            countElements += 1;
                                            dictionaryOfElementNumbers.Add(guidString, countElements);
                                            dictionaryOfElementsQuantity.Add(guidString, 1);
                                        }
                                        //добавление мультивыносок
                                        СreatingMLeaders(objEntity, dictionaryOfElementNumbers, guidString, countElements);
                                    }
                                    else
                                    {
                                        editor.WriteMessage("\nGUID не найден.");
                                    }
                                }
                                else
                                {
                                    editor.WriteMessage("\nДля данного элемента не найдена доп. информация.");
                                }
                            }
                        }
                    }
                }
                // Завершение транзакции и вызов метода отрисовки таблицы
                transaction.Commit();
                CreateTable(countElements, dictionaryOfElementNumbers, dictionaryOfElementsQuantity);
            }
        }


        //метод для поиска GUID в свойстве xData сущности 
        private string GetGuidFromXData(ResultBuffer xData)
        {
            TypedValue[] values = xData.AsArray();
            foreach (TypedValue value in values)//перебираем данные из xData
            {
                if (value.TypeCode == 1000) // 1000- код типа данных, которые представляют собой строку ASCII (длиной до 255 байт) в xData.               
                {
                    string potentialGuid = value.Value.ToString();
                    Guid guid;
                    if (Guid.TryParse(potentialGuid, out guid))// // здесь метод TryParse проверяет, можно ли получить 128-битовое целое число из строки
                    {
                        return potentialGuid;
                    }
                }
            }
            return null;
        }

        //метод создания и отрисовки мультивыносок
        public void СreatingMLeaders(Entity objEntity, Dictionary<string, int> dictionaryOfElementNumbers, string guidString, int countElements)
        {
            Point3d startPoint = objEntity.GeometricExtents.MinPoint;//получение координат крайней точки выбранного элемента
            startPoint = startPoint.Add(new Vector3d(10, 10, 0)); // смещение точки на 10 единиц
            Point3d endPoint = startPoint.Add(new Vector3d(10, -100, 0)); // смещение конечной точки
           //создание мультивыноски и текста, установка свойств
            MLeader leader = new MLeader();
            leader.SetDatabaseDefaults();
            MText mText = new MText();
            mText.SetDatabaseDefaults();
            mText.TextHeight = 20.0;
            mText.SetContentsRtf(dictionaryOfElementNumbers[guidString].ToString());
            mText.Location = endPoint;
            leader.MText = mText;
            // Установка начальной и конечной точек
            int idx = leader.AddLeaderLine(startPoint);
            Transaction transaction = document.TransactionManager.StartTransaction();
            using (transaction)
            {
                BlockTable bt = (BlockTable)transaction.GetObject(dataBase.BlockTableId, OpenMode.ForRead);
                BlockTableRecord modelSpace = (BlockTableRecord)transaction.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                modelSpace.AppendEntity(leader);
                transaction.AddNewlyCreatedDBObject(leader, true);
                transaction.Commit();
            }
        }

        //метод создания таблицы и отображения ее на чертеже
        internal void CreateTable(int countOfRows, Dictionary<string, int> listOfNumbersElements, Dictionary<string, int> quantityDictionary)
        {
            try
            {
                DatabaseSqliteClass databaseSqliteClass = new DatabaseSqliteClass();//создание экземпляра класса DatabaseSqliteClass для работы с БД SQLite
                int k = 2;
                PromptPointResult pr =editor.GetPoint("\nВыберете место(точку) для отоюражения таблицы : ");
                if (pr.Status == PromptStatus.OK)
                {
                    //создание таблицы и установка ее свойств 
                    Table table = new Table();
                    table.TableStyle = dataBase.Tablestyle;
                    table.NumRows = countOfRows + 2;
                    table.NumColumns = 4;
                    table.SetRowHeight(40);
                    table.SetColumnWidth(350);
                    table.Position = pr.Value;
                    //перебор ячеек таблицы для их форматирования
                    for (int i = 0; i < countOfRows + 2; i++)
                    {
                        for (int j = 0; j < 4; j++)
                        {
                            table.SetTextHeight(i, j, 15);
                            table.SetAlignment(i, j, CellAlignment.MiddleCenter);
                        }
                    }

                    // установка шапки таблицы
                    table.SetTextString(1, 0, "№");
                    table.SetTextString(1, 1, "Наименование");
                    table.SetTextString(1, 2, "Количество");
                    table.SetTextString(1, 3, "Вес");
               
                 //установка в ячейки таблицы значений из словаря и БД SQLite
                    foreach (var number in listOfNumbersElements)
                    {
                        System.Data.DataTable dataTable = databaseSqliteClass.getInfoOfElememts(number.Key);
                        table.SetTextString(k, 0, number.Value.ToString());
                        table.SetTextString(k, 1, dataTable.Rows[0]["Name"].ToString());
                        table.SetTextString(k, 3, dataTable.Rows[0]["Weight"].ToString());
                        k++;
                    }
                    k = 2;

                    //установка в ячейки таблицы значений из словаря
                    foreach (var quantity in quantityDictionary)
                    {
                        table.SetTextString(k, 2, quantity.Value.ToString());
                        k++;
                    }

                    table.GenerateLayout();
                    //добавление таблицы в пространство модели
                    Transaction transaction = document.TransactionManager.StartTransaction();
                    using (transaction)
                    {
                        BlockTable bt = (BlockTable)transaction.GetObject(document.Database.BlockTableId, OpenMode.ForRead);
                        BlockTableRecord btr = (BlockTableRecord)transaction.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                        btr.AppendEntity(table);
                        transaction.AddNewlyCreatedDBObject(table, true);
                        transaction.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }


        //метод удаления доп.информации с чертежа
        public void EraseBlocks(Database db)
        {
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                // Удаление объектов из модели пространства
                foreach (ObjectId objId in modelSpace)
                {
                    Entity entity = tr.GetObject(objId, OpenMode.ForWrite) as Entity;
                    if (entity.GetType()==typeof(Table) || entity.GetType() == typeof(MLeader)) //проверка типов блоков, чтобы не удалить блоки изначального чертежа
                        entity.Erase();
                }
                tr.Commit();
            }
        }

    }
}
