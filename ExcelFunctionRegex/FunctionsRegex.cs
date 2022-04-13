using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelFunctionRegex
{
    public class FunctionsRegex
    {
        /// <summary>
        /// Функция поиска по регулярному выражению
        /// </summary>
        /// <param name="text">текст в котором производится поиск по регулярному выражению</param>
        /// <param name="pattern">шаблон поиска по регулярному выражению</param>
        /// <param name="item">номер группы найденного результата</param>
        /// <returns></returns>
        [ExcelFunction(Description = "Поиск по регулярному выражению", Name = "ПоискПоРегВыр")]
        public static string RegExpExtract([ExcelArgument(Description = "Текст в котором производиться поиск", Name = "Текст")] string text,
            [ExcelArgument(Description = "Текст шаблона регулярного выражения", Name = "Регулярное выражение")] string pattern,
            [ExcelArgument(Description = "Номер результата в группе поиска по регулярному выражению", Name = "Номер результата")] int numItem,
          [ExcelArgument(Description = "Номер группы результата поиска по регулярному выражению", Name = "Номер группы")] int numGroup) {

            try {
                // Произведём поиск по регулярному выражению.
                MatchCollection matches = (new Regex(pattern)).Matches(text);
                // Если что-то было найдено,
                if (matches.Count > 0) {
                    // то проверим правильно ли указан индекс поиска результата.
                    if (numGroup < 0 || numItem < 0)
                        throw new ArgumentOutOfRangeException();
                    GroupCollection groups = matches[numItem].Groups;
                    // Если номер группы указан больше чем количество групп, то так же вернём исключение.
                    if (numGroup > groups.Count - 1)
                        throw new ArgumentOutOfRangeException();
                    return groups[numGroup].Value;
                }
                return "#СОВПАДЕНИЙ_НЕ_НАЙДЕНО";
            } catch (ArgumentOutOfRangeException) {
                return "#НЕТ_ДАННЫХ";
            } catch (ArgumentException) {
                return "#ОШИБКА_ШАБЛОНА_РЕГУЛЯРНОГО_ВЫРАЖЕНИЯ";
            } catch(Exception) {
                return "#ОШИБКА";
            }
        }

        /// <summary>
        /// Функция разделение текста по регулярному выражению
        /// </summary>
        /// <param name="text"></param>
        /// <param name="pattern"></param>
        /// <param name="item"></param>
        /// <returns></returns>
        [ExcelFunction(Description = "Разделение текста по регулярному выражению", Name = "ДелТекстПоРегВыр")]
        public static string RegExpSplit([ExcelArgument(Description = "Текст который требуется разделить", Name = "Текст")] string text,
            [ExcelArgument(Description = "Текст шаблона регулярного выражения", Name = "Регулярное выражение")] string pattern,
            [ExcelArgument(Description = "Номер значения массива полученных значений", Name = "Номер значения")] int item) {
            try {
                Regex regex = new Regex(pattern);
                return regex.Split(text)[item];
            } catch (IndexOutOfRangeException) {
                return "#НЕТ_ДАННЫХ";
            } catch (ArgumentException) {
                return "#ОШИБКА_ШАБЛОНА_РЕГУЛЯРНОГО_ВЫРАЖЕНИЯ";
            }
        }

        /// <summary>
        /// Функция проверки по регулярному выражению
        /// </summary>
        /// <param name="text">текст в котором производится поиск по регулярному выражению</param>
        /// <param name="pattern">шаблон поиска по регулярному выражению</param>
        /// <returns></returns>
        [ExcelFunction(Description = "Проверка по регулярному выражению", Name = "ПроверкаПоРегВыр")]
        public static object RegExpCheck([ExcelArgument(Description = "Текст в котором производиться поиск", Name = "Текст")] string text,
            [ExcelArgument(Description = "Текст шаблона регулярного выражения", Name = "Регулярное выражение")] string pattern) {
            try {
                Regex regex = new Regex(pattern);
                return regex.IsMatch(text);
            } catch (ArgumentException) {
                return "#ОШИБКА_ШАБЛОНА_РЕГУЛЯРНОГО_ВЫРАЖЕНИЯ";
            }
        }

        /// <summary>
        /// Объединение текста через знак разделитель
        /// </summary>
        /// <param name="range"></param>
        /// <param name="separator"></param>
        /// <returns></returns>
        [ExcelFunction(Description = "Объединение текста через знак разделитель", Name = "СцепТекстПоРазд")]
        public static string JoinText([ExcelArgument(Description = "Диапазон ячеек для объединения", Name = "Текст")] object[] range,
            [ExcelArgument(Description = "Текст разделитель списка", Name = "разделитель")] string separator) {
            string result = "";
            // Пробежимся по переданным объектам,
            for (int i=0; i< range.Length; i++) {
                // и если этот объект не пустой,
                if (range[i] != ExcelEmpty.Value)
                    // то добавим его в общую строку с переданным разделителем.
                    result += $"{range[i]}{separator}";
            }
            // Если строка была сформированна, то венем ее без послднего разделитея. А если нет, то и вернем пустую строку.
            return result.Length > separator.Length ? result.Substring(0, result.Length - separator.Length) : "";
        }
    }
}
