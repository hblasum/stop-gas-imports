Writeup: Full-Russian-gas-embargo-with-securing-European-industry.docx

Eurostat custom tables archived for reproducibility:

* custom-table-imports-of-natural-gas-by-partner-country.xlsx -> Eurostat. “Custom Dataset: Imports of Natural Gas by Partner Country.” Accessed April 4, 2022. https://ec.europa.eu/eurostat/databrowser/view/NRG_TI_GAS__custom_2428849/default/table?lang=en. This custom dataset is based on Eurostat. “Imports of Natural Gas by Partner Country - Products Datasets - Eurostat.” Accessed April 4, 2022. https://ec.europa.eu/eurostat/web/products-datasets/-/nrg_ti_gas)
* custom-table-supply-transformation-and-consumption-of-gas.xlsx -> Eurostat: Custom Table: Supply, Transformation and Consumption of Gas.” Accessed March 30, 2022. https://ec.europa.eu/eurostat/databrowser/view/NRG_CB_GAS__custom_2395063/default/table?lang=en  .This custom table is based on Eurostat. “Statistics | Eurostat: Supply, Transformation and Consumption of Gas.” Accessed March 30, 2022. https://ec.europa.eu/eurostat/databrowser/view/NRG_CB_GAS/default/table?lang=en&category=nrg.nrg_quant.nrg_quanta.nrg_cb
* mapping-of-fine-granular-data.xlsx: derived from custom-table-supply-transformation-and-consumption-of-gas.xlsx
* consumer-savings-calculation.xlsx: calculation of consumer savings 

Helper scripts for table analysis:

* eurostat_custom_table_supply_transformation_and_consumption_of_gas_analyse.py # script doing the aggregation of data
* eurostat.py # initial helper script
