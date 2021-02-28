from Check_files import CheckFiles, Distributors


checks = CheckFiles()
distributors = Distributors(checks.paths['Distributors'])
debtors = distributors.debtors
# print(debtors)
