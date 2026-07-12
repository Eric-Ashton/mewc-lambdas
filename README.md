# MEWC Lambda Library
A version-controlled library of custom Excel LAMBDA functions and VBA for MS Excel World Championship (MEWC). I use these to solve the puzzles in MEWC faster and more accurately than I could without them. 

The latest version of the template can be found at https://tinyurl.com/EricLambas and using this repo is not necessary to try them out. If you want, you can pick or choose code out of the repo to use in your own work. 

## Contents
| Path | What |
|------|------|
| `lambdas/*.lambda` | One file per lambda: signature, comment, code, description. |
| `vba/*.bas` | VBA modules — repo sync (`repo_sync`), unit-test + lambda-management tools (`unit_test_tools`), utilities (`utils`), plus the workbook's supporting subs. |
| `tools/lambda_check.py` | Authoring-rule checker (run before committing). |
| `MEWC Lambda and VBA Unit Tests.xlsm` | The committed test workbook — all lambdas, VBA, Prep sheet, and unit-test sheets. |
| `CONVENTIONS.md` | Authoring rules + repo/Excel workflow (single source of truth). |
| `MEWC Lambdas Edit and Test Workflow.md` | The edit → test → review → merge loop. |

## License
Released under the [MIT License](LICENSE) — free to use, adapt, and share.
Attribution is appreciated but not required.

## Acknowledgements
Thank you to the MS Excel World Championship / Excel esports community members who
shared code, helped, and inspired this library:

- **Diarmuid Early** — [YouTube](https://www.youtube.com/@DimEarly) · [Lambda video](https://www.youtube.com/watch?v=257yIalp5DU)
- **Hayden Wiseman** — [YouTube](https://www.youtube.com/@ExcelFinance-j2p)
- **Brittany Deaton** — [Lambda video](https://www.youtube.com/watch?v=UEtRNvZDCD8)
- **Lorenzo Foti** — [YouTube](https://www.youtube.com/@LorenzoFoti)
- **Jeremy Freelove** — [YouTube](https://www.youtube.com/@ExcelGladiator-365)
- **Julian Poeltl** — [YouTube](https://www.youtube.com/@excelwithExcel-xlsx)
- **Bo Rydobon** — [YouTube](https://www.youtube.com/@ExcelWizard/videos)
- **Juan José Cifuentes** — [YouTube](https://www.youtube.com/@excelmanchile)
- **Coby Dombowsky**
- …and anyone else I forgot!

### About Excel esports
- <https://excel-esports.com/>
- [The New York Times — Microsoft Excel World Championships](https://www.nytimes.com/2025/01/20/us/microsoft-excel-world-championships.html)
- [The Wall Street Journal — Microsoft World Excel Championships](https://www.wsj.com/tech/microsoft-world-excel-championships-las-vegas-448c5f0b)
   