/*
 * ------------------------------------------------------------------------
 *
 *  Copyright by KNIME AG, Zurich, Switzerland
 *  Website: http://www.knime.com; Email: contact@knime.com
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
 *  KNIME and ECLIPSE being a combined program, KNIME AG herewith grants
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
 * ---------------------------------------------------------------------
 *
 * History
 *   Nov 3, 2020 (Simon Schmid, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader.read;

import static java.util.stream.Collectors.toSet;

import java.io.IOException;
import java.io.InputStream;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader.SheetIterator;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.util.CheckUtils;
import org.knime.core.util.UniqueNameGenerator;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.AreaOfSheetToRead;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderConfig;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelCell.KNIMECellType;
import org.knime.filehandling.core.node.table.reader.config.TableReadConfig;
import org.knime.filehandling.core.node.table.reader.read.IntervalRead;
import org.knime.filehandling.core.node.table.reader.read.Read;
import org.knime.filehandling.core.node.table.reader.spec.TypedReaderColumnSpec;
import org.knime.filehandling.core.node.table.reader.spec.TypedReaderTableSpec;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

/**
 * Utility class for Excel nodes.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
public final class ExcelUtils {

    private static final Pattern EXCEL_COLUMN_NAME_REGEX = Pattern.compile("[A-Z]+");

    private ExcelUtils() {
        // Hide constructor, utils class
    }

    /**
     * Assigns names to the columns in {@link TypedReaderTableSpec spec} if they don't contain a name already. The
     * naming scheme is A, B, C, etc. whereby the letter depends on the column index in the original table, i.e., the
     * first missing column name will not necessarily be assigned 'A' but the letter that corresponds with its index.
     *
     * @param spec {@link TypedReaderTableSpec} containing columns to assign names if they are missing
     * @param config the config
     * @param hiddenColumns the set of hidden columns
     * @return a {@link TypedReaderTableSpec} with the same types as {@link TypedReaderTableSpec spec} in which all
     *         columns are named
     */
    public static TypedReaderTableSpec<KNIMECellType> assignNamesIfMissing(
        final TypedReaderTableSpec<KNIMECellType> spec, final TableReadConfig<ExcelTableReaderConfig> config,
        final Set<Integer> hiddenColumns) {
        final UniqueNameGenerator nameGen = new UniqueNameGenerator(spec.stream()//
            .map(TypedReaderColumnSpec::getName)//
            .map(n -> n.orElse(null))//
            .filter(Objects::nonNull)//
            .collect(toSet()));
        return new TypedReaderTableSpec<>(IntStream.range(0, spec.size())
            .mapToObj(i -> assignNameIfMissing(i, spec.getColumnSpec(i), nameGen, config, hiddenColumns))
            .collect(Collectors.toList()));
    }

    private static <T> TypedReaderColumnSpec<T> assignNameIfMissing(final int idx, final TypedReaderColumnSpec<T> spec,
        final UniqueNameGenerator nameGen, final TableReadConfig<ExcelTableReaderConfig> config,
        final Set<Integer> hiddenColumns) {
        final Optional<String> name = spec.getName();
        if (name.isPresent()) {
            return spec;
        } else {
            return TypedReaderColumnSpec.createWithName(
                nameGen.newName(ExcelUtils.getExcelColumnName(getFilteredIdx(idx, config, hiddenColumns))),
                spec.getType(), spec.hasType());
        }
    }

    private static int getFilteredIdx(final int idx, final TableReadConfig<ExcelTableReaderConfig> config,
        final Set<Integer> hiddenColumns) {
        int filteredIdx;
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        final int startColIdx;
        // get the start index
        if (excelConfig.getAreaOfSheetToRead() == AreaOfSheetToRead.PARTIAL) {
            startColIdx = getColIntervalStart(config);
            filteredIdx = idx + startColIdx;
        } else {
            startColIdx = 0;
            filteredIdx = idx;
        }
        final int rowIdColIdx = getRowIdColIdx(excelConfig.getRowIDCol());
        if (excelConfig.isSkipHiddenCols()) {
            // increment the index by the number of hidden columns that are before it and >= start
            final int i = filteredIdx;
            int count = (int)hiddenColumns.stream().filter(hIdx -> hIdx >= startColIdx && hIdx <= i).count();
            filteredIdx = filteredIdx + count;
            // increment as long as we have an index that is hidden
            while (hiddenColumns.contains(filteredIdx)) {
                filteredIdx++;
            }
            // increment if we had a row ID before the index and >= start (as this was filtered out and not counted)
            // make sure the row ID was not hidden as we have already covered that case above
            if (rowIdBeforeIdxAndWithinStart(config.useRowIDIdx(), filteredIdx, startColIdx, rowIdColIdx)
                && !hiddenColumns.contains(rowIdColIdx)) {
                filteredIdx++;
            }
        } else if (!excelConfig.isUseRawSettings()
            && rowIdBeforeIdxAndWithinStart(config.useRowIDIdx(), filteredIdx, startColIdx, rowIdColIdx)) {
            // increment if we had a row ID before the index and >= start (as this was filtered out and not counted)
            filteredIdx++;
        }
        return filteredIdx;
    }

    private static boolean rowIdBeforeIdxAndWithinStart(final boolean useRowIdIdx, final int filteredIdx,
        final int startColIdx, final int rowIdColIdx) {
        return useRowIdIdx && rowIdColIdx <= filteredIdx && rowIdColIdx >= startColIdx;
    }

    private static int getColIntervalStart(final TableReadConfig<ExcelTableReaderConfig> config) {
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        if (excelConfig.getAreaOfSheetToRead() == AreaOfSheetToRead.ENTIRE) {
            return 0;
        }
        return ExcelUtils.getFirstColumnIdx(excelConfig.getReadFromCol());
    }

    /**
     * Returns the name of the first sheet with data, i.e., non-empty sheet. If there is no sheet with data, the first
     * one is returned.
     *
     * @param sheetNames the map of sheet names and booleans that indicate if a sheet is the first with data
     *
     * @return the name of the first sheet with data or first sheet if all empty
     */
    public static String getFirstSheetWithDataOrFirstIfAllEmpty(final Map<String, Boolean> sheetNames) {
        if (sheetNames.isEmpty()) {
            return ""; // should never happen; still, handle gracefully to prevent UI crashes
        }
        final String firstSheet = sheetNames.keySet().iterator().next();
        for (final Entry<String, Boolean> sheetNameEntry : sheetNames.entrySet()) {
            if (sheetNameEntry.getValue().booleanValue()) {
                return sheetNameEntry.getKey();
            }
        }
        return firstSheet;
    }

    /**
     * Returns a map that contains the names of the sheets contained in the specified {@link Workbook} as keys and
     * whether it is the first non-empty sheet as value.
     *
     * @param workbook the workbook
     * @return the map of sheet names and whether a sheet is the first with data
     */
    public static Map<String, Boolean> getSheetNames(final Workbook workbook) {
        final Map<String, Boolean> sheetNames = new LinkedHashMap<>(); // LinkedHashMap to retain order
        boolean nonEmptySheetFound = false;
        for (final Sheet sheet : workbook) {
            if (nonEmptySheetFound) {
                sheetNames.put(sheet.getSheetName(), false);
            } else {
                final boolean isEmpty = isEmpty(sheet);
                nonEmptySheetFound = !isEmpty;
                sheetNames.put(sheet.getSheetName(), !isEmpty);
            }
        }
        return sheetNames;
    }

    /**
     * Returns a map that contains the names of the sheets contained in the file read by the specified
     * {@link XSSFReader} as keys and whether it is the first non-empty sheet as value.
     *
     * @param xmlReader the xml reader
     * @param reader the xssf reader
     * @param sharedStrings the shared strings table
     * @return the map of sheet names and whether a sheet is the first with data
     * @throws InvalidFormatException
     * @throws IOException
     * @throws SAXException
     */
    public static Map<String, Boolean> getSheetNames(final XMLReader xmlReader, final XSSFReader reader,
        final SharedStrings sharedStrings) throws InvalidFormatException, IOException, SAXException {
        final Map<String, Boolean> sheetNames = new LinkedHashMap<>(); // LinkedHashMap to retain order
        boolean nonEmptySheetFound = false;
        xmlReader.setContentHandler(
            new XSSFSheetXMLHandler(reader.getStylesTable(), sharedStrings, new IsEmpty(), new DataFormatter(), false));
        final SheetIterator sheetsData = (SheetIterator)reader.getSheetsData();
        while (sheetsData.hasNext()) {
            try (final InputStream inputStream = sheetsData.next()) {
                if (nonEmptySheetFound) {
                    sheetNames.put(sheetsData.getSheetName(), false);
                } else {
                    final boolean sheetEmpty = isSheetEmpty(xmlReader, inputStream);
                    sheetNames.put(sheetsData.getSheetName(), !sheetEmpty);
                    nonEmptySheetFound = !sheetEmpty;
                }
            }
        }
        return sheetNames;
    }

    /**
     * Returns the Excel column name for the given column index.
     *
     * @param colIdx the index
     * @return the Excel column name or {@code null} if {@code colIdx} is < 0
     */
    public static String getExcelColumnName(final int colIdx) {
        if (colIdx < 0) {
            return null;
        }
        if (colIdx < 26) {
            return String.valueOf((char)('A' + colIdx));
        }
        final int mod = colIdx % 26;
        final int div = colIdx / 26;
        return getExcelColumnName(div - 1) + String.valueOf((char)('A' + mod));
    }

    /**
     * Parses the column label and returns its index. While the label is one-based, the returned index is zero-based.
     *
     * @param columnLabel the column label to parse, either numeric or capital letters
     * @return the column index
     */
    public static int getColumnIndex(final String columnLabel) {
        if (columnLabel.isEmpty()) {
            return -1;
        }
        CheckUtils.checkArgument(EXCEL_COLUMN_NAME_REGEX.matcher(columnLabel).matches(), "Invalid column '%s'.",
            columnLabel);

        int result = 0;
        for (int d = 0; d < columnLabel.length(); d++) {
            final int dig = toNumber(columnLabel.charAt(columnLabel.length() - d - 1));
            result += dig * (int)Math.round(Math.pow(26, d));
        }
        return result - 1;
    }

    private static int toNumber(final char c) {
        assert c <= 'Z';
        assert c > '@';
        return c - '@';
    }

    /**
     * Parses the row label and returns its index. While the label is one-based, the returned index is zero-based. If
     * the value is empty, the default value will be returned.
     *
     * @param rowLabel the row label to parse
     * @param defaultValue the default value
     * @return the parsed row index
     */
    public static int rowNumberStringToIdx(final String rowLabel, final int defaultValue) {
        final String trimmedCol = rowLabel.trim();
        int rowIdx = defaultValue;
        if (!trimmedCol.isEmpty()) {
            try {
                rowIdx = Integer.parseInt(trimmedCol) - 1;
            } catch (NumberFormatException e) {
                throw new IllegalArgumentException("Invalid row number '" + rowLabel + "'.");
            }
        }
        CheckUtils.checkArgument(rowIdx >= 0, "Invalid row number '%s'. Specify a number >= 1.", rowLabel);
        return rowIdx;
    }

    /**
     * Validates the first and last row.
     *
     * @param firstRow first row
     * @param lastRow last row
     * @throws InvalidSettingsException if last row < first row
     */
    public static void validateRowIndexes(final String firstRow, final String lastRow) throws InvalidSettingsException {
        CheckUtils.checkSetting(rowNumberStringToIdx(firstRow, 0) <= rowNumberStringToIdx(lastRow, Integer.MAX_VALUE),
            "The last row must not be before the first row.");
    }

    /**
     * Validates the first and last column.
     *
     * @param firstCol first column
     * @param lastCol last column
     * @throws InvalidSettingsException if last column < first column
     */
    public static void validateColIndexes(final String firstCol, final String lastCol) throws InvalidSettingsException {
        final int lastColumnIdx = getLastColumnIdx(lastCol);
        CheckUtils.checkSetting(lastColumnIdx < 0 || getFirstColumnIdx(firstCol) <= lastColumnIdx,
            "The last column must not be before the first column.");
    }

    private static int columnToIdx(final String columnString, final int defaultValue) {
        final String trimmedCol = columnString.trim();
        int colIdx = defaultValue;
        if (!trimmedCol.isEmpty()) {
            try {
                colIdx = Integer.parseInt(trimmedCol) - 1;
            } catch (NumberFormatException e) {
                colIdx = ExcelUtils.getColumnIndex(trimmedCol);
            }
        }
        return colIdx;
    }

    /**
     * Returns the index of the row ID column. If the string is empty or invalid, an {@code IllegalArgumentException} is
     * thrown.
     *
     * @param string the input string
     * @return the row ID index
     */
    public static int getRowIdColIdx(final String string) {
        int rowIdColumn = columnToIdx(string, -1);
        CheckUtils.checkArgument(rowIdColumn >= 0,
            "The row ID column is invalid. It must be a number >= 1 or a name starting with A.");
        return rowIdColumn;
    }

    /**
     * Returns the index of the row ID column. If the string is invalid, an {@code IllegalArgumentException} is thrown.
     * If the string is empty, 0 is returned.
     *
     * @param string the input string
     * @return the index of the first included column
     */
    public static int getFirstColumnIdx(final String string) {
        int firstColumn = columnToIdx(string, 0);
        CheckUtils.checkArgument(firstColumn >= 0,
            "The first column is invalid. It must be a number >= 1 or a name starting with A.");
        return firstColumn;
    }

    /**
     * Returns the index of the row ID column. If the string is invalid, an {@code IllegalArgumentException} is thrown.
     * If the string is empty, -1 is returned.
     *
     * @param string the input string
     * @return the index of the last included column
     */
    public static int getLastColumnIdx(final String string) {
        return columnToIdx(string, -1);
    }

    /**
     * Decorate the {@link Read} with row filtering {@link Read}s.
     *
     * @param read the {@link Read} to decorate
     * @param config the config
     * @return the decorated {@link Read}
     */
    public static Read<ExcelCell> decorateRowFilterReads(Read<ExcelCell> read,
        final TableReadConfig<ExcelTableReaderConfig> config) {
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        if (excelConfig.getAreaOfSheetToRead() == AreaOfSheetToRead.PARTIAL) {
            read = new IntervalRead<>(read, getFirstIncludedRowIdx(config), getLastIncludedRowIdx(config) + 1L);
        }
        if (config.skipEmptyRows()) {
            read = new ExcelSkipEmptyRead<>(read, config.useRowIDIdx());
        }
        if (excelConfig.isSkipHiddenRows()) {
            read = new SkipHiddenRowsRead(read);
        }
        return read;
    }

    private static int getFirstIncludedRowIdx(final TableReadConfig<ExcelTableReaderConfig> config) {
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        final int startRowIdx = ExcelUtils.rowNumberStringToIdx(excelConfig.getReadFromRow(), 0);
        // we need to move the start one further up if there is the column header row before
        if (config.useColumnHeaderIdx() && config.getColumnHeaderIdx() < startRowIdx) {
            return startRowIdx - 1;
        }
        return startRowIdx;
    }

    private static int getLastIncludedRowIdx(final TableReadConfig<ExcelTableReaderConfig> config) {
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        final int endRowIdx = ExcelUtils.rowNumberStringToIdx(excelConfig.getReadToRow(), Integer.MAX_VALUE - 1);
        // we need to move the end one further up if there is the column header row before or the end is the header
        if (config.useColumnHeaderIdx() && config.getColumnHeaderIdx() <= endRowIdx) {
            return endRowIdx - 1;
        }
        return endRowIdx;
    }

    private static boolean isSheetEmpty(final XMLReader xmlReader, final InputStream inputStream)
        throws IOException, SAXException {
        try {
            xmlReader.parse(new InputSource(inputStream));
            return true;
        } catch (ParsingInterruptedException e) { // NOSONAR, exception is expected and handled
            // not empty
            return false;
        }
    }

    /**
     * We need to iterate over all rows and columns as a cell could be blank but have a style set. In such cases,
     * {@link Sheet#getLastRowNum()} would not return 0 and a row would not be {@code null}.
     */
    private static boolean isEmpty(final Sheet sheet) {
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            final Row row = sheet.getRow(i);
            if (row != null && !isRowEmpty(row)) {
                return false;
            }
        }
        return true;
    }

    private static boolean isRowEmpty(final Row row) {
        for (int j = 0; j < row.getLastCellNum(); j++) {
            final Cell cell = row.getCell(j, MissingCellPolicy.RETURN_BLANK_AS_NULL);
            if (cell != null) {
                return false;
            }
        }
        return true;
    }

    /**
     * Checks whether the sheet is empty or not.
     */
    static class IsEmpty implements SheetContentsHandler {

        @Override
        public void cell(final String cellReference, final String formattedValue, final XSSFComment comment) {
            throw new ParsingInterruptedException();
        }

        @Override
        public void startRow(final int rowNum) {
            // do nothing
        }

        @Override
        public void endRow(final int rowNum) {
            // do nothing
        }

        @Override
        public void headerFooter(final String text, final boolean isHeader, final String tagName) {
            // do nothing
        }

    }

    /**
     * Exception to be thrown when the thread that parses the sheet should be interrupted.
     *
     * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
     */
    static class ParsingInterruptedException extends RuntimeException {

        private static final long serialVersionUID = 1L;

    }

}
