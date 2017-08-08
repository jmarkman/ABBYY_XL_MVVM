using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ABBYY_XL_MVVM.Components
{
    public class WorkstationRow
    {
        private List<string> dataRow;
        public int RowLength { get; set; }

        /// <summary>
        /// Constructor. Initializes row for item storage.
        /// </summary>
        public WorkstationRow()
        {
            dataRow = new List<string>();
            RowLength = dataRow.Count;
        }

        /// <summary>
        /// Wrapper for List.Add(). Adds the content of the cell to the column.
        /// </summary>
        /// <param name="value">The cell content as a string</param>
        public void Add(string value)
        {
            dataRow.Add(value);
            RowLength++;
        }

        /// <summary>
        /// Wrapper for List.Remove(). Removes a piece of content from the column.
        /// </summary>
        /// <param name="value">The content to remove from the column as a string</param>
        public void Remove(string value)
        {
            dataRow.Remove(value);
            RowLength--;
        }

        /// <summary>
        /// Wrapper for List.RemoveAt(). Removes a piece of content from the column.
        /// </summary>
        /// <param name="index">The content to remove from the column at a specified index</param>
        public void RemoveFromRow(int index)
        {
            dataRow.RemoveAt(index);
            RowLength--;
        }

        /// <summary>
        /// Replace the value at the specified index with another
        /// </summary>
        /// <param name="index">The numerical position where the replacement should occur in the list</param>
        /// <param name="value">The value to provide as replacement</param>
        public void ReplaceAt(int index, string value)
        {
            dataRow[index] = value;
        }

        /// <summary>
        /// Wrapper for index-based access of List. Returns the item at the specified index.
        /// </summary>
        /// <param name="index">Location to access as an integer</param>
        /// <returns>The data at the specified index</returns>
        public string DataAtIndex(int index)
        {
            return dataRow[index];
        }
    }
}
