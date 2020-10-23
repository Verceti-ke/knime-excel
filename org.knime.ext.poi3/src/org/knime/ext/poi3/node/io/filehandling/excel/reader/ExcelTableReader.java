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
 *   Oct 13, 2020 (Simon Schmid, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node.io.filehandling.excel.reader;

import java.io.IOException;
import java.nio.file.Path;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Predicate;

import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.knime.core.node.ExecutionMonitor;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelRead;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.XLSRead;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.XLSXRead;
import org.knime.filehandling.core.node.table.reader.TableReader;
import org.knime.filehandling.core.node.table.reader.config.TableReadConfig;
import org.knime.filehandling.core.node.table.reader.read.Read;
import org.knime.filehandling.core.node.table.reader.spec.TableSpecGuesser;
import org.knime.filehandling.core.node.table.reader.spec.TypedReaderTableSpec;
import org.knime.filehandling.core.node.table.reader.type.hierarchy.TreeTypeHierarchy;
import org.knime.filehandling.core.node.table.reader.type.hierarchy.TypeFocusableTypeHierarchy;
import org.knime.filehandling.core.node.table.reader.type.hierarchy.TypeTester;

/**
 * Reader for Excel files.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
final class ExcelTableReader implements TableReader<ExcelTableReaderConfig, Class<?>, String> {

    static final TypeFocusableTypeHierarchy<Class<?>, String> TYPE_HIERARCHY = createHierarchy();

    @Override
    public Read<String> read(final Path path, final TableReadConfig<ExcelTableReaderConfig> config) throws IOException {
        return getExcelRead(path, config);
    }

    @Override
    public TypedReaderTableSpec<Class<?>> readSpec(final Path path,
        final TableReadConfig<ExcelTableReaderConfig> config, final ExecutionMonitor exec) throws IOException {
        final TableSpecGuesser<Class<?>, String> guesser = createGuesser();
        try (final ExcelRead read = getExcelRead(path, config)) {
            return guesser.guessSpec(read, config, exec);
        }
    }

    private static ExcelRead getExcelRead(final Path path, final TableReadConfig<ExcelTableReaderConfig> config)
        throws IOException {
        try {
            final String pathLowerCase = path.toString().toLowerCase();
            if (pathLowerCase.endsWith(".xlsx") || pathLowerCase.endsWith(".xlsm")) {
                return new XLSXRead(path, config);
            }
        } catch (OLE2NotOfficeXmlFileException e) { // NOSONAR
            // Happens if an xls file has been specified that ends with xlsx or xlsm.
            // We do not fail but simply use the XLSParser instead.
        }
        return new XLSRead(path, config);
    }

    private static TableSpecGuesser<Class<?>, String> createGuesser() {
        return new TableSpecGuesser<>(createHierarchy(), Function.identity());
    }

    // TODO copied from CSV reader for now, need to be adjusted for Excel types
    private static TypeFocusableTypeHierarchy<Class<?>, String> createHierarchy() {
        return TreeTypeHierarchy.builder(createTypeTester(String.class, t -> {
        })).addType(String.class, createTypeTester(Double.class, Double::parseDouble))
            .addType(Double.class, createTypeTester(Long.class, Long::parseLong))
            .addType(Long.class, createTypeTester(Integer.class, Integer::parseInt)).build();
    }

    private static TypeTester<Class<?>, String> createTypeTester(final Class<?> type, final Consumer<String> tester) {
        return TypeTester.createTypeTester(type, consumerToPredicate(tester));
    }

    private static Predicate<String> consumerToPredicate(final Consumer<String> tester) {
        return s -> {
            try {
                tester.accept(s);
                return true;
            } catch (NumberFormatException ex) {
                return false;
            }
        };
    }

}
