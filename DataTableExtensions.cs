using System.ComponentModel;
using System.Data;
using System.Diagnostics.Contracts;
using System.Reflection;
using System.Text;

namespace Loxifi
{
    /// <summary>
    ///
    /// </summary>
    public static class DataTableExtensions
    {
        /// <summary>
        /// Adds an object to the data table as a data row, and then returns the new row
        /// </summary>
        /// <param name="dt">The data table to target</param>
        /// <param name="toAdd">The object to add to the table</param>
        /// <returns>The newly created data row containing the object information</returns>
        public static DataRow Add(this DataTable dt, object toAdd)
        {
            if (dt is null)
            {
                throw new ArgumentNullException(nameof(dt));
            }

            if (toAdd is null)
            {
                throw new ArgumentNullException(nameof(toAdd));
            }

            if (toAdd is DataRow dr)
            {
                dt.Rows.Add(dr);
                return dr;
            }

            DataRow newRow = dt.NewRow();

            foreach (PropertyInfo pi in toAdd.GetType().GetProperties())
            {
                if (pi.GetGetMethod() != null)
                {
                    object val = pi.GetValue(toAdd);

                    newRow[pi.Name] = val is null ? DBNull.Value : val;
                }
            }

            dt.Rows.Add(newRow);

            return newRow;
        }

        /// <summary>
        /// Adds an item to the data row, adding the column if needed.
        /// If the column exists, updates the existing value
        /// </summary>
        /// <param name="dr">The data row to update</param>
        /// <param name="column">The column to target</param>
        /// <param name="value">The value to add or update</param>
        public static void AddOrUpdate(this DataRow dr, string column, object value)
        {
            if (dr is null)
            {
                throw new ArgumentNullException(nameof(dr));
            }

            _ = dr.Table.EnsureColumn(column);
            dr[column] = value;
        }

        /// <summary>
        /// Returns true if the data table contains the requested column
        /// </summary>
        /// <param name="dt">The data table to check</param>
        /// <param name="columnName">The column to check for</param>
        /// <param name="comparison">An optional string comparison to use when checking for the column name. Defaults to OrdinalIgnoreCase</param>
        /// <returns>True if a column with a matching name is found</returns>
        public static bool ContainsColumn(this DataTable dt, string columnName, StringComparison comparison = StringComparison.OrdinalIgnoreCase)
        {
            return dt is null
                ? throw new ArgumentNullException(nameof(dt))
                : string.IsNullOrEmpty(columnName)
                ? throw new ArgumentException($"'{nameof(columnName)}' cannot be null or empty.", nameof(columnName))
                : dt.Columns.Cast<DataColumn>().Any(dc => string.Equals(columnName, dc.ColumnName, comparison));
        }

        /// <summary>
        /// Creates a column on a table if it doesn't already exist
        /// </summary>
        /// <param name="dt">The data table to target</param>
        /// <param name="columnName">The name of the column to ensure</param>
        /// <param name="columnType">The type of the data contained in the column</param>
        /// <returns>True if the column already existed. False if it was created</returns>
        public static bool EnsureColumn(this DataTable dt, string columnName, Type? columnType = null)
        {
            if (dt is null)
            {
                throw new ArgumentNullException(nameof(dt));
            }

            if (string.IsNullOrEmpty(columnName))
            {
                throw new ArgumentException($"'{nameof(columnName)}' cannot be null or empty.", nameof(columnName));
            }

            columnType ??= typeof(string);

            if (columnType.IsGenericType && columnType.GetGenericTypeDefinition() == typeof(Nullable<>))
            {
                columnType = Nullable.GetUnderlyingType(columnType);
            }

            if (dt.ContainsColumn(columnName))
            {
                return true;
            }

            _ = dt.Columns.Add(columnName, columnType);

            return false;
        }

        /// <summary>
        /// Fills a datatable with an IEnumerable of objects, with property names as headers and values as items
        /// </summary>
        /// <typeparam name="TData">The base type of the ienumerable that will be used to scaffold the table</typeparam>
        /// <param name="thisTable">The table to fill</param>
        /// <param name="data">The item IEnumerable to add</param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void Fill<TData>(this DataTable thisTable, IEnumerable<TData> data)
        {
            if (thisTable is null)
            {
                throw new ArgumentNullException(nameof(thisTable));
            }

            if (data is null)
            {
                throw new ArgumentNullException(nameof(data));
            }

            List<PropertyInfo> Properties = new();

            Dictionary<PropertyInfo, int> PropertyOrder = new();

            foreach (PropertyInfo thisProp in typeof(TData).GetProperties().Reverse())
            {
                int index = 0;

                while (index < Properties.Count && PropertyOrder[Properties.ElementAt(index)] > index)
                {
                    index++;
                }

                Properties.Insert(index, thisProp);
                PropertyOrder.Add(thisProp, index);
            }

            foreach (PropertyInfo thisProperty in Properties)
            {
                DisplayNameAttribute displayNameAttribute = thisProperty.GetCustomAttribute<DisplayNameAttribute>();
                string DisplayName = displayNameAttribute != null ? displayNameAttribute.DisplayName : thisProperty.Name;
                _ = thisTable.Columns.Add(DisplayName);
            }

            foreach (TData thisObj in data)
            {
                DataRow thisRow = thisTable.NewRow();
                int i = 0;
                foreach (PropertyInfo thisProperty in Properties)
                {
                    thisRow[i++] = thisProperty.GetValue(thisObj);
                }

                thisTable.Rows.Add(thisRow);
            }
        }

        /// <summary>
        /// Sets up a data table to have columns representative of properties found on a given type
        /// </summary>
        /// <param name="dt">The data table to scaffold</param>
        /// <param name="toScaffold">The type to use to create the columns on the table</param>
        public static void Scaffold(this DataTable dt, Type toScaffold)
        {
            if (dt is null)
            {
                throw new ArgumentNullException(nameof(dt));
            }

            if (toScaffold is null)
            {
                throw new ArgumentNullException(nameof(toScaffold));
            }

            foreach (PropertyInfo pi in toScaffold.GetProperties())
            {
                if (pi.GetGetMethod() != null)
                {
                    _ = dt.EnsureColumn(pi.Name, pi.PropertyType);
                }
            }
        }

        /// <summary>
        /// Returns an HTML string table representation of a data table
        /// </summary>
        /// <param name="dt">The data table to convert</param>
        /// <returns>An HTML string table representation of a data table</returns>
        public static string ToHtmlTable(this DataTable dt)
        {
            Contract.Requires(dt != null);

            StringBuilder body = new();

            _ = body.Append("<table><tr>");

            foreach (DataColumn dc in dt.Columns)
            {
                _ = body.Append($"<th>{dc.ColumnName}</th>");
            }

            _ = body.Append("</tr>");

            foreach (DataRow dr in dt.Rows)
            {
                _ = body.Append("<tr>");

                foreach (object o in dr.ItemArray)
                {
                    _ = body.Append($"<td>{o}</td>");
                }

                _ = body.Append("</tr>");
            }

            _ = body.Append("</table>");

            return body.ToString();
        }
    }
}