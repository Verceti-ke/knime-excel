/*
 * ------------------------------------------------------------------------
 *
 *  Copyright (C) 2003 - 2010
 *  University of Konstanz, Germany and
 *  KNIME GmbH, Konstanz, Germany
 *  Website: http://www.knime.org; Email: contact@knime.org
 *
 *  This program is free software; you can redistribute it and/or modify
 *  it under the terms of the GNU General Public License, Version 3, as
 *  published by the Free Software Foundation.
 *
 *  This program is distributed in the hope that it will be useful, but
 *  WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 *  GNU General Public License for more details.
 *
 *  You should have received a copy of the GNU General Public License
 *  along with this program; if not, see <http://www.gnu.org/licenses>.
 *
 *  Additional permission under GNU GPL version 3 section 7:
 *
 *  KNIME interoperates with ECLIPSE solely via ECLIPSE's plug-in APIs.
 *  Hence, KNIME and ECLIPSE are both independent programs and are not
 *  derived from each other. Should, however, the interpretation of the
 *  GNU GPL Version 3 ("License") under any applicable laws result in
 *  KNIME and ECLIPSE being a combined program, KNIME GMBH herewith grants
 *  you the additional permission to use and propagate KNIME together with
 *  ECLIPSE with only the license terms in place for ECLIPSE applying to
 *  ECLIPSE and the GNU GPL Version 3 applying for KNIME, provided the
 *  license terms of ECLIPSE themselves allow for the respective use and
 *  propagation of ECLIPSE together with KNIME.
 *
 *  Additional permission relating to nodes for KNIME that extend the Node
 *  Extension (and in particular that are based on subclasses of NodeModel,
 *  NodeDialog, and NodeView) and that only interoperate with KNIME through
 *  standard APIs ("Nodes"):
 *  Nodes are deemed to be separate and independent programs and to not be
 *  covered works.  Notwithstanding anything to the contrary in the
 *  License, the License does not apply to Nodes, you are not required to
 *  license Nodes under the License, and you are granted a license to
 *  prepare and propagate Nodes, in each case even if such Nodes are
 *  propagated with or for interoperation with KNIME.  The owner of a Node
 *  may freely choose the license terms applicable to such Node, including
 *  when such Node is propagated with or for interoperation with KNIME.
 * -------------------------------------------------------------------
 *
 * History
 *   Jun 25, 2009 (ohl): created
 */
package org.knime.ext.poi.node.read2;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashSet;
import java.util.Set;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.knime.core.data.DataCell;
import org.knime.core.data.DataColumnSpec;
import org.knime.core.data.DataColumnSpecCreator;
import org.knime.core.data.DataRow;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.DataType;
import org.knime.core.data.RowIterator;
import org.knime.core.data.def.DoubleCell;
import org.knime.core.data.def.IntCell;
import org.knime.core.data.def.StringCell;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeLogger;

/**
 *
 * @author Peter Ohl, KNIME.com, Zurich, Switzerland
 */
public class XLSTableSettings {

    private static final NodeLogger LOGGER =
            NodeLogger.getLogger(XLSTableSettings.class);

    private final XLSUserSettings m_userSettings;

    private final DataTableSpec m_spec;

    private final Set<Integer> m_skippedCols;

    /**
     *
     */
    public XLSTableSettings(final XLSUserSettings userSettings)
            throws InvalidSettingsException, FileNotFoundException,
            IOException, InvalidFormatException {

        String errMsg = userSettings.getStatus(false);
        if (errMsg != null) {
            throw new IllegalArgumentException(errMsg);
        }

        m_userSettings = XLSUserSettings.clone(userSettings);

        // set the range of the rows and columns to read
        if (m_userSettings.getReadAllData()) {
            // clear any possibly set first/last row/col
            m_userSettings.setFirstColumn(-1);
            m_userSettings.setLastColumn(-1);
            m_userSettings.setFirstRow(-1);
            m_userSettings.setLastRow(-1);
            m_userSettings.setReadAllData(false);
        }
        if (m_userSettings.getLastColumn() < 0 ||
                m_userSettings.getLastRow() < 0) {
            // if bounds are not user set - figure them out
            setMinMaxColumnAndRow(m_userSettings);
        }

        // this analyzes the types of the columns
        HashSet<Integer> skippedcols = new HashSet<Integer>();
        m_spec = createSpec(m_userSettings, skippedcols);
        m_skippedCols = new HashSet<Integer>();
        if (skippedcols != null) {
            m_skippedCols.addAll(skippedcols);
        }
    }

    /**
     * Uses the specified table spec and skipped cols to get the column names
     * and column types from. They must fit the file set in the user settings!
     * This constructor does not analyze the XL file.
     *
     */
    private XLSTableSettings(final XLSUserSettings userSettings,
            final DataTableSpec tableSpec, final Set<Integer> skippedCols)
            throws InvalidSettingsException, FileNotFoundException, IOException {
        String errMsg = userSettings.getStatus(false);
        if (errMsg != null) {
            throw new IllegalArgumentException(errMsg);
        }
        if (userSettings.getReadAllData()) {
            throw new IllegalArgumentException("Constructor not suitable if"
                    + " area of interest is not set.");
        }

        m_userSettings = XLSUserSettings.clone(userSettings);

        // this analyzes the types of the columns
        m_skippedCols = new HashSet<Integer>();
        if (skippedCols != null) {
            m_skippedCols.addAll(skippedCols);
        }
        m_spec = tableSpec;

    }

    /**
     *
     * @return
     */
    public DataTableSpec getDataTableSpec() {
        return m_spec;
    }

    /**
     * Runs through the selected sheet and sets the minimum of all first cell
     * and the maximum of all last cell numbers of all rows. That is, it returns
     * the columns containing data in the data sheet. (If the minimum is already
     * set to a valid index, it doesn't change it - in order to support for open
     * ranges.) It ignores all "skip" flags, i.e. count in empty and hidden
     * columns.
     *
     * @param settings specifies the sheet to read. firstColumn and lastColumn
     *            in this object will be set by this method
     * @throws InvalidSettingsException
     * @throws IOException
     * @throws FileNotFoundException
     * @throws InvalidFormatException
     */
    private static void setMinMaxColumnAndRow(final XLSUserSettings settings)
            throws InvalidSettingsException, IOException,
            FileNotFoundException, InvalidFormatException {
        if (settings == null) {
            throw new NullPointerException("Settings can't be null");
        }
        if (settings.getFileLocation() == null) {
            throw new NullPointerException("File location must be set.");
        }
        if (settings.getSheetName() == null
                || settings.getSheetName().isEmpty()) {
            throw new InvalidSettingsException("Sheet name must be set.");

        }

        FileInputStream fs = new FileInputStream(settings.getFileLocation());
        BufferedInputStream inp = new BufferedInputStream(fs);

        inp = new BufferedInputStream(fs);
        Workbook wb = WorkbookFactory.create(inp);

        Sheet sheet = wb.getSheet(settings.getSheetName());
        if (sheet == null) {
            throw new InvalidSettingsException("Sheet '"
                    + settings.getSheetName() + "' not available");
        }

        short firstColIdx = -1; // stores the cell index
        short lastColNum = -1; // stores the cell index + 1

        int maxRowIdx = XLSTable.getLastRowIdx(sheet);
        int minRowIdx = sheet.getFirstRowNum();

        for (int row = minRowIdx; row <= maxRowIdx; row++) {
            Row r = sheet.getRow(row);
            if (r == null) {
                continue;
            }
            if (r.getFirstCellNum() < 0) {
                // no cell in this row
                continue;
            }
            if (firstColIdx < 0 || r.getFirstCellNum() < firstColIdx) {
                firstColIdx = r.getFirstCellNum();
            }
            if (lastColNum < r.getLastCellNum()) {
                lastColNum = r.getLastCellNum();
            }
        }

        lastColNum--; // now it's the index

        if (firstColIdx < 0 || lastColNum < 0 || lastColNum < firstColIdx) {
            // if first col is set, don't change it
            if (settings.getFirstColumn() < 0) {
                settings.setFirstColumn(0);
                settings.setLastColumn(0);
            } else {
                settings.setLastColumn(settings.getFirstColumn());
            }
        } else {
            // if first col is set, don't change it
            if (settings.getFirstColumn() < 0) {
                settings.setFirstColumn(firstColIdx);
                settings.setLastColumn(lastColNum);
            } else {
                int last = lastColNum;
                if (last < settings.getFirstColumn()) {
                    settings.setLastColumn(settings.getFirstColumn());
                } else {
                    settings.setLastColumn(last);
                }
            }
        }

        if (settings.getFirstRow() < 0) {
            settings.setFirstRow(minRowIdx);
            settings.setLastRow(maxRowIdx);
        } else {
            int last = maxRowIdx;
            if (last < settings.getLastRow()) {
                settings.setLastRow(settings.getFirstRow());
            } else {
                settings.setLastRow(last);
            }
        }

        fs.close();
    }

    /**
     * Creates and returns a new table spec from the current settings.
     *
     * @return a new table spec from the current settings
     * @throws InvalidSettingsException if settings are invalid.
     * @throws InvalidFormatException
     */
    private static DataTableSpec createSpec(final XLSUserSettings settings,
            final Set<Integer> skippedCols) throws InvalidSettingsException,
            IOException, InvalidFormatException {

        ArrayList<DataType> columnTypes =
                analyzeColumnTypes(settings, skippedCols);

        int numOfCols = columnTypes.size();
        String[] colHdrs = createColHeaders(settings, numOfCols, skippedCols);
        assert colHdrs.length == numOfCols;

        DataColumnSpec[] colSpecs = new DataColumnSpec[numOfCols];

        for (int col = 0; col < numOfCols; col++) {
            assert (col < columnTypes.size() && columnTypes.get(col) != null);
            colSpecs[col] =
                    new DataColumnSpecCreator(colHdrs[col], columnTypes
                            .get(col)).createSpec();
        }

        // create a name
        String tableName =
                "XL table: " + new File(settings.getFileLocation()).getName()
                        + "[" + settings.getSheetName() + "]";
        return new DataTableSpec(tableName, colSpecs);
    }

    /**
     * Either reads the column names from the sheet, of generates new ones.
     */
    private static String[] createColHeaders(final XLSUserSettings settings,
            final int numOfCols, final Set<Integer> skippedCols)
            throws InvalidSettingsException {

        String[] result = null;
        if (settings.getHasColHeaders() && !settings.getKeepXLColNames()) {
            result = readColumnHeaders(settings, numOfCols, skippedCols);
        }
        if (result == null) {
            result = new String[numOfCols];
        }
        fillEmptyColHeaders(settings, skippedCols, result);
        return result;
    }

    private static void fillEmptyColHeaders(final XLSUserSettings settings,
            final Set<Integer> skippedCols, final String[] colHdrs) {
        // XL Sheets don't have more than 256 columns
        HashSet<String> names = new HashSet<String>();
        for (int i = 0; i < colHdrs.length; i++) {
            if (colHdrs[i] != null) {
                names.add(colHdrs[i]);
            }
        }

        int xlOffset = 0;

        for (int i = 0; i < colHdrs.length; i++) {
            while (skippedCols.contains(xlOffset)) {
                xlOffset++;
            }
            if (colHdrs[i] == null) {
                String colName = "Col" + i;
                if (settings.getKeepXLColNames()) {
                    colName =
                            XLSTable.getColLabel(settings.getFirstColumn()
                                    + xlOffset);
                }
                colHdrs[i] = getUniqueName(colName, names);
                names.add(colHdrs[i]);
            }
            xlOffset++;
        }
    }

    /**
     * Looks at the specified XLS file and tries to determine the type of all
     * columns contained. Also returns the number of columns and rows in all
     * sheets.
     *
     * @param settings the pre-set settings.
     * @return the result settings
     * @throws IOException if the specified file couldn't be read.
     * @throws InvalidSettingsException if settings are invalid
     * @throws InvalidFormatException
     */
    private static ArrayList<DataType> analyzeColumnTypes(
            final XLSUserSettings settings, final Set<Integer> skippedCols)
            throws IOException, InvalidSettingsException,
            InvalidFormatException {

        if (settings == null) {
            throw new NullPointerException("Settings can't be null");
        }
        if (settings.getFileLocation() == null) {
            throw new NullPointerException("File location must be set.");
        }

        FileInputStream fs = new FileInputStream(settings.getFileLocation());
        BufferedInputStream inp = new BufferedInputStream(fs);

        try {
            inp = new BufferedInputStream(fs);
            Workbook wb = WorkbookFactory.create(inp);
            return setColumnTypes(wb, settings, skippedCols);

        } finally {
            if (fs != null) {
                fs.close();
            }
        }

    }

    /**
     * Traverses the specified sheet in the file and detects the type for all
     * columns in the sheets.
     *
     * @param settings
     * @param resultSettings
     * @param skipEmtpyCols
     * @param skipHiddenCols
     * @throws IOException
     */
    private static ArrayList<DataType> setColumnTypes(final Workbook wb,
            final XLSUserSettings settings, final Set<Integer> skippedCols)
            throws IOException {

        int colNum = settings.getLastColumn() - settings.getFirstColumn() + 1;

        ArrayList<DataType> colTypes =
                new ArrayList<DataType>(Arrays.asList(new DataType[colNum]));
        skippedCols.clear();

        String dbgMsg = "";
        Sheet sh = wb.getSheet(settings.getSheetName());
        if (sh != null) {
            int maxRowIdx = XLSTable.getLastRowIdx(sh);
            int minRowIdx = sh.getFirstRowNum();
            if (settings.getLastRow() < maxRowIdx) {
                maxRowIdx = settings.getLastRow();
            }
            if (settings.getFirstRow() > minRowIdx) {
                minRowIdx = settings.getFirstRow();
            }
            for (int row = minRowIdx; row <= maxRowIdx; row++) {
                Row r = sh.getRow(row);
                if (r == null) {
                    continue;
                }
                // skip the row that contains the column names
                if (settings.getHasColHeaders()
                        && row == settings.getColHdrRow()) {
                    continue;
                }
                int knimeColIdx = -1;
                for (int xlCol = settings.getFirstColumn(); xlCol <= settings
                        .getLastColumn(); xlCol++) {
                    knimeColIdx++;
                    if (settings.getHasRowHeaders()
                            && xlCol == settings.getRowHdrCol()) {
                        // skip the column with the row IDs
                        skippedCols.add(xlCol);
                        continue;
                    }
                    if (sh.isColumnHidden(xlCol)
                            && settings.getSkipHiddenColumns()) {
                        skippedCols.add(xlCol);
                        continue;
                    }
                    Cell cell = r.getCell(xlCol);
                    dbgMsg =
                            "Cell ("
                                    + xlCol
                                    + ","
                                    + row
                                    + ")"
                                    + " KNIME Col"
                                    + knimeColIdx
                                    + ", POI type="
                                    + (cell != null ? getPoiTypeName(cell
                                            .getCellType()) : ": <null> ")
                                    + ", ";
                    if (cell != null) {
                        // determine the type
                        switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_BLANK:
                            // missing cell - doesn't change any type
                            dbgMsg += " <missing>";
                            break;
                        case Cell.CELL_TYPE_BOOLEAN:
                            // KNIME has no boolean - use String
                            colTypes.set(knimeColIdx, StringCell.TYPE);
                            dbgMsg += " KNIME type: String";
                            break;
                        case Cell.CELL_TYPE_ERROR:
                            // treated as missing cell
                            dbgMsg += " <missing>";
                            break;
                        case Cell.CELL_TYPE_FORMULA:
                            DataType currType = colTypes.get(knimeColIdx);
                            if (currType == null
                                    || IntCell.TYPE.equals(currType)) {
                                colTypes.set(knimeColIdx, DoubleCell.TYPE);
                                dbgMsg += " KNIME type: Formula/Double";
                            } else {
                                dbgMsg += " KNIME type - unchanged";
                            }
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            // numeric could be double, int or date
                            if (colTypes.get(knimeColIdx) == StringCell.TYPE) {
                                // string takes all
                                dbgMsg += " KNIME type - unchanged";
                                break;
                            }
                            if (DateUtil.isCellDateFormatted(cell)) {
                                // we use StringCells for date format
                                // (using a DataAndTime cell leads to UTC time.
                                // With string, we get the entered time (as
                                // string) and it can be translated in date/time
                                // with an additional node)
                                colTypes.set(knimeColIdx, StringCell.TYPE);
                                dbgMsg += " KNIME type: String (for date/time)";
                                break;
                            }
                            Double num = cell.getNumericCellValue();
                            if (num.isInfinite() || num.isNaN()) {
                                // kind of a missing value
                                break;
                            }
                            if (new Double(num.intValue()).equals(num)) {
                                // could be represented as int
                                if (colTypes.get(knimeColIdx) == null) {
                                    colTypes.set(knimeColIdx, IntCell.TYPE);
                                    dbgMsg += " KNIME type: Int";
                                } else {
                                    // it should be double - which is fine.
                                    dbgMsg += " KNIME type - unchanged";
                                }
                            } else {
                                colTypes.set(knimeColIdx, DoubleCell.TYPE);
                                dbgMsg += " KNIME type: Double";
                            }
                            break;
                        case Cell.CELL_TYPE_STRING:
                            colTypes.set(knimeColIdx, StringCell.TYPE);
                            dbgMsg += " KNIME type: String";
                            break;
                        default:
                            LOGGER.error("Unexpected cell type ("
                                    + cell.getCellType() + ")");
                            dbgMsg += " KNIME type - unchanged";
                            break;
                        }

                    }
                    LOGGER.debug(dbgMsg);
                }

            }
        }

        // null types represent empty columns (except skipped hidden cols)

        dbgMsg = "";
        for (int c = 0; c < colTypes.size(); c++) {
            int xlCol = c + settings.getFirstColumn();
            DataType type = colTypes.get(c);
            if (skippedCols.contains(xlCol)) {
                if (settings.getHasRowHeaders()
                        && xlCol == settings.getRowHdrCol()) {
                    dbgMsg += "Col" + c + ": <rowHdrCol:removed> ";
                } else {
                    dbgMsg += "Col" + c + ": <hidden:removed> ";
                }
                continue;
            } else if (type == null) {
                if (settings.getSkipEmptyColumns()) {
                    dbgMsg += "Col" + c + ": <empty:removed> ";
                    skippedCols.add(xlCol);
                    continue;
                } else {
                    dbgMsg += "Col" + c + ": <empty:DataCell> ";
                    colTypes.set(c, DataType.getType(DataCell.class));
                    continue;
                }
            } else {
                dbgMsg += "Col" + c + ": " + type.toString() + "  ";
                continue;
            }
        }
        LOGGER.debug(dbgMsg);
        // remove skipped cols from type list
        if (skippedCols.size() > 0) {
            ArrayList<DataType> result = new ArrayList<DataType>();
            for (int c = 0; c < colTypes.size(); c++) {
                if (!skippedCols.contains(c + settings.getFirstColumn())) {
                    // (skippedCols contains xl-indices)
                    result.add(colTypes.get(c));
                }
            }
            return result;
        } else {
            return colTypes;
        }
    }

    private static String getPoiTypeName(final int poiType) {
        switch (poiType) {
        case Cell.CELL_TYPE_BLANK:
            return "blank";
        case Cell.CELL_TYPE_STRING:
            return "string";
        case Cell.CELL_TYPE_BOOLEAN:
            return "boolean";
        case Cell.CELL_TYPE_ERROR:
            return "error";
        case Cell.CELL_TYPE_NUMERIC:
            return "numeric";
        case Cell.CELL_TYPE_FORMULA:
            return "formula";
        }
        return "unknown[" + poiType + "]";
    }

    /**
     * Reads the contents of the specified column header row. Names not
     * specified in the sheet are null in the result. Uniquifies the names from
     * the sheet. The length of the returned array is determined by the first
     * and last column to read. The result is null, if the hasColHdr flag is not
     * set, or the specified header row is not in the sheet.
     *
     * @param settings valid settings
     * @param numOfCols number of columns to read headers for
     * @param skippedCols columns in the sheet being skipped (due to emptiness
     *            or row headers)
     * @return Returns null, if sheet contains no column headers, or the
     *         specified header row is not in the sheet, otherwise an array with
     *         unique names, or null strings, if the sheet didn't contain a name
     *         for that column.
     * @throws InvalidSettingsException if settings are invalid
     */
    private static String[] readColumnHeaders(final XLSUserSettings settings,
            final int numOfCols, final Set<Integer> skippedCols)
            throws InvalidSettingsException {

        String[] result = new String[numOfCols];

        if (!settings.getHasColHeaders()) {
            return result;
        }

        try {

            XLSUserSettings colHdrSettings = XLSUserSettings.clone(settings);

            // this avoids endless recursion
            colHdrSettings.setKeepXLNames(true);

            // analyze the file and set the table spec
            XLSTableSettings tableSettings =
                    new XLSTableSettings(colHdrSettings);

            // now use the settings to read the header row
            colHdrSettings.setHasColHeaders(false);
            // the row ID refers to the row in the XL file
            colHdrSettings.setSkipEmptyRows(false);
            // the table settings contain the interesting range
            colHdrSettings.setFirstColumn(tableSettings.getFirstColumn());
            colHdrSettings.setLastColumn(tableSettings.getLastColumn());
            // we only need the header row
            colHdrSettings.setFirstRow(colHdrSettings.getColHdrRow());
            colHdrSettings.setLastRow(colHdrSettings.getColHdrRow());

            // create a table spec for reading col headers (all string cols)
            DataColumnSpec[] colSpecs =
                    new DataColumnSpec[tableSettings.m_spec.getNumColumns()];
            for (int i = 0; i < colSpecs.length; i++) {
                colSpecs[i] =
                        new DataColumnSpecCreator("Col" + i, StringCell.TYPE)
                                .createSpec();
            }
            // create the table settings to read the header row

            XLSTableSettings tableS =
                    new XLSTableSettings(colHdrSettings, new DataTableSpec(
                            colSpecs), tableSettings.m_skippedCols);
            XLSTable table = new XLSTable(tableS);
            RowIterator iterator = table.iterator();
            DataRow row = null;
            if (iterator.hasNext()) {
                row = iterator.next();
            }
            if (row == null) {
                // table had too few rows
                LOGGER.warn("Specified column header row not contained "
                        + "in sheet");
                return result;
            }
            if (numOfCols != row.getNumCells()) {
                LOGGER.debug("The column header row just read has "
                        + row.getNumCells() + " cells, but we expected "
                        + numOfCols);
            }
            HashSet<String> names = new HashSet<String>();
            for (int i = 0; i < row.getNumCells(); i++) {
                if (i >= result.length) {
                    break;
                }
                if (!row.getCell(i).isMissing()) {
                    result[i] = getUniqueName(row.getCell(i).toString(), names);
                    names.add(result[i]);
                }
            }

        } catch (Exception e) {
            NodeLogger.getLogger(XLSUserSettings.class).debug(
                    "Caught exception while accessing file "
                            + settings.getFileLocation()
                            + ". Creating synthetic column names", e);
        }

        return result;
    }

    private static String getUniqueName(final String name,
            final Set<String> names) {
        int cnt = 2;
        String unique = name;
        while (names.contains(unique)) {
            unique = name + "_" + cnt++;
        }
        return unique;
    }

    /*
     * ---------------- setter and getter ------------------------------------
     */

    /**
     * @return the fileLocation
     */
    public String getFileLocation() {
        return m_userSettings.getFileLocation();
    }

    /**
     * @return the index of the sheet to read
     */
    public String getSheetName() {
        return m_userSettings.getSheetName();
    }

    /**
     * @return the firstRow
     */
    public int getFirstRow() {
        return m_userSettings.getFirstRow();
    }

    /**
     * @return the lastRow
     */
    public int getLastRow() {
        return m_userSettings.getLastRow();
    }

    /**
     * @return the firstColumn
     */
    public int getFirstColumn() {
        return m_userSettings.getFirstColumn();
    }

    /**
     * @return the lastColumn
     */
    public int getLastColumn() {
        return m_userSettings.getLastColumn();
    }

    /**
     * Set with empty or hidden columns. The index in the set is the offset from
     * the firstColumn.
     *
     * @return the skippedCols
     */
    public Set<Integer> getSkippedCols() {
        return Collections.unmodifiableSet(m_skippedCols);
    }

    /**
     * @return the skipEmptyRows
     */
    public boolean getSkipEmptyRows() {
        return m_userSettings.getSkipEmptyRows();
    }

    /**
     * @return the colHdrRow
     */
    public int getColHdrRow() {
        return m_userSettings.getColHdrRow();
    }

    /**
     * @return the hasColHeaders
     */
    public boolean getHasColHeaders() {
        return m_userSettings.getHasColHeaders();
    }

    /**
     * @return the hasRowHeaders
     */
    public boolean getHasRowHeaders() {
        return m_userSettings.getHasRowHeaders();
    }

    /**
     * @return the rowHdrCol
     */
    public int getRowHdrCol() {
        return m_userSettings.getRowHdrCol();
    }

    /**
     * @return the missValuePattern
     */
    public String getMissValuePattern() {
        return m_userSettings.getMissValuePattern();
    }

    /**
     * @return
     * @see org.knime.ext.poi.node.read2.XLSUserSettings#getKeepXLColNames()
     */
    public boolean getKeepXLColNames() {
        return m_userSettings.getKeepXLColNames();
    }

    /**
     * @return
     * @see org.knime.ext.poi.node.read2.XLSUserSettings#getUniquifyRowIDs()
     */
    public boolean getUniquifyRowIDs() {
        return m_userSettings.getUniquifyRowIDs();
    }

}