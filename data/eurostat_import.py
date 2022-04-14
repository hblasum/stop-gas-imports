import openpyxl

path = "."

wbs = openpyxl.load_workbook(filename=f"{path}/custom-table-supply-transformation-and-consumption-of-gas.xlsx", data_only=True,
                                read_only=True)
wbi = openpyxl.load_workbook(filename=f"{path}/custom-table-imports-of-natural-gas-by-partner-country.xlsx", data_only=True,
                                read_only=True)

household_savings_rate = 0.35
transformation_savings_rate = 0.97
industry_savings_rate = 0.20

print(f'Country\tHouseholds\tCommercial\tElectricity generation\tIndustry savings\tSubstitutable gas\tIndustry use\tRussian imports\tImports needed for industry\tPossible exports to other countries')

sum_households = 0
sum_commercial = 0
sum_transformation_use = 0
sum_industry_use = 0
sum_gas_imports_russia = 0
sum_gas_imports_needed_for_industry = 0
sum_possible_gas_exports = 0
sum_substitutable_gas = 0
sum_industry_savings = 0

for country in range(15, 41):
    commercial = 0.0
    households = 0.0
    transformation_use = 0.0
    industry_use = 0.0
    industry_savings = 0.0
    division = 1000.0
    for ws in wbs.worksheets:
        if ws.title == "Sheet 1":
            country_name = ws[f"a{country}"].value
            if country_name == "Germany (until 1990 former territory of the FRG)":
                    country_name = "Germany"
            print(country_name, end="")
        if ws[f"k{country}"].value is not None and not isinstance(ws[f"k{country}"].value, str) and ws.title in ['Sheet 21', 'Sheet 22', 'Sheet 23', 'Sheet 25', 'Sheet 35', 'Sheet 36', 'Sheet 37', 'Sheet 39', 'Sheet 42', 'Sheet 43', 'Sheet 45', 'Sheet 48', 'Sheet 60', 'Sheet 62', 'Sheet 63', 'Sheet 64', 'Sheet 66', 'Sheet 68', 'Sheet 70', 'Sheet 71', 'Sheet 72', 'Sheet 74', 'Sheet 76', 'Sheet 78', 'Sheet 80', 'Sheet 82', 'Sheet 91', 'Sheet 100', 'Sheet 102']:
            industry_use += ws[f"k{country}"].value
        if ws[f"k{country}"].value is not None and ws.title in ['Sheet 20', 'Sheet 96']:
            households += ws[f"k{country}"].value
        if ws[f"k{country}"].value is not None and ws.title in ['Sheet 98']:
            commercial += ws[f"k{country}"].value
        if ws[f"k{country}"].value is not None and ws.title in ['Sheet 18', 'Sheet 19']:
            transformation_use += ws[f"k{country}"].value
        industry_savings = industry_savings_rate * industry_use
        substitutable_gas = (households + commercial) * household_savings_rate + transformation_use * transformation_savings_rate + \
            industry_savings
        gas_imports_russia = wbi["Sheet 1"][f"k{country}"].value + wbi["Sheet 3"][f"k{country}"].value
        gas_imports_needed_for_industry = 0
        possible_gas_exports = 0
        if gas_imports_russia > substitutable_gas:
                gas_imports_needed_for_industry = gas_imports_russia - substitutable_gas
        else:
                possible_gas_exports = substitutable_gas - gas_imports_russia
    print(f'\t{round(households / division * household_savings_rate)}\t{round(commercial / division * household_savings_rate)}'
          f'\t{round(transformation_use * transformation_savings_rate / division)}'
          f'\t{round(industry_savings/division)}'
          f'\t{round(substitutable_gas/division)}'
          f'\t{round((industry_use - industry_savings)/division)}'
          f'\t{round(gas_imports_russia/division)}'
          f'\t{round(gas_imports_needed_for_industry/division)}\t{round(possible_gas_exports/division)}')

    sum_households += households/division
    sum_commercial += commercial/division
    sum_transformation_use += transformation_use/division
    sum_industry_use += industry_use/division
    sum_gas_imports_russia += gas_imports_russia/division
    sum_gas_imports_needed_for_industry += gas_imports_needed_for_industry/division
    sum_possible_gas_exports += possible_gas_exports/division
    sum_substitutable_gas += substitutable_gas/division
    sum_industry_savings += industry_savings/division

print(f'SUM\t{round(sum_households)}\t{round(sum_commercial)}\t{round(sum_transformation_use)}'
      f'\t{round(sum_industry_savings)}'
      f'\t{round(sum_substitutable_gas)}'
      f'\t{round(sum_industry_use)}'
      f'\t{round(sum_gas_imports_russia)}'
          f'\t{round(sum_gas_imports_needed_for_industry)}\t{round(sum_possible_gas_exports)}')
