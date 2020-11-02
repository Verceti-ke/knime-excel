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

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettings;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.NotConfigurableException;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.filehandling.core.node.table.reader.config.ConfigSerializer;
import org.knime.filehandling.core.node.table.reader.config.DefaultMultiTableReadConfig;
import org.knime.filehandling.core.node.table.reader.config.DefaultTableReadConfig;

/**
 * TODO implement once dialog is extended with more settings
 *
 * {@link ConfigSerializer} for the Excel reader node.
 *
 * @author Simon Schmid, KNIME GmbH, Konstanz, Germany
 */
enum ExcelMultiTableReadConfigSerializer implements
    ConfigSerializer<DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>>> {

        /**
         * Singleton instance.
         */
        INSTANCE;

    private static final String CFG_ADVANCED_SETTINGS_TAB = "advanced_settings";

    private static final String CGF_USE_15_DIGITS_PRECISION = "use_15_digits_precision";

    @Override
    public void loadInDialog(
        final DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>> config,
        final NodeSettingsRO settings, final PortObjectSpec[] specs) throws NotConfigurableException {
        loadAdvancedSettingsTabInDialog(config, getOrEmpty(settings, CFG_ADVANCED_SETTINGS_TAB));
    }

    @Override
    public void loadInModel(
        final DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>> config,
        final NodeSettingsRO settings) throws InvalidSettingsException {
        loadAdvancedSettingsTabInModel(config, settings.getNodeSettings(CFG_ADVANCED_SETTINGS_TAB));
        // TODO add settings to dialog, for now hard-coded
        final DefaultTableReadConfig<?> tc = config.getTableReadConfig();
        tc.setAllowShortRows(true);
        tc.setUseColumnHeaderIdx(true);
        tc.setColumnHeaderIdx(0);
        tc.setLimitRowsForSpec(false);
    }

    @Override
    public void saveInModel(
        final DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>> config,
        final NodeSettingsWO settings) {
        saveAdvancedSettingsTab(config, settings.addNodeSettings(CFG_ADVANCED_SETTINGS_TAB));
    }

    @Override
    public void validate(final NodeSettingsRO settings) throws InvalidSettingsException {
        validateAdvancedSettingsTab(settings.getNodeSettings(CFG_ADVANCED_SETTINGS_TAB));
    }

    @Override
    public void saveInDialog(
        final DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>> config,
        final NodeSettingsWO settings) throws InvalidSettingsException {
        saveAdvancedSettingsTab(config, settings.addNodeSettings(CFG_ADVANCED_SETTINGS_TAB));
    }

    private static void loadAdvancedSettingsTabInDialog(
        final DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>> config,
        final NodeSettingsRO settings) {
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        excelConfig.setUse15DigitsPrecision(settings.getBoolean(CGF_USE_15_DIGITS_PRECISION, true));
    }

    private static void loadAdvancedSettingsTabInModel(
        final DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>> config,
        final NodeSettingsRO settings) throws InvalidSettingsException {
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        excelConfig.setUse15DigitsPrecision(settings.getBoolean(CGF_USE_15_DIGITS_PRECISION));
    }

    private static void saveAdvancedSettingsTab(
        final DefaultMultiTableReadConfig<ExcelTableReaderConfig, DefaultTableReadConfig<ExcelTableReaderConfig>> config,
        final NodeSettingsWO settings) {
        final ExcelTableReaderConfig excelConfig = config.getReaderSpecificConfig();
        settings.addBoolean(CGF_USE_15_DIGITS_PRECISION, excelConfig.isUse15DigitsPrecision());
    }

    public static void validateAdvancedSettingsTab(final NodeSettingsRO settings) throws InvalidSettingsException {
        settings.getBoolean(CGF_USE_15_DIGITS_PRECISION);
    }

    private static NodeSettingsRO getOrEmpty(final NodeSettingsRO settings, final String key) {
        try {
            return settings.getNodeSettings(key);
        } catch (InvalidSettingsException ise) { // NOSONAR
            return new NodeSettings(key);
        }
    }

}
