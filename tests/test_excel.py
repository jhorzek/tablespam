from tablespam.Excel._as_excel.create_test_files import create_test_files_cars
import openpyxl


def test_excel(tmp_path):
    test_xlsx = create_test_files_cars()

    # The created xlsx files correspond to the excel files in data.
    # We load those files and compare our results in test_xlsx to them
    for tst in test_xlsx:
        # because openpyxl may not read the workbook the same it was written, we
        # will first first save our new workbook and then read it again.
        test_xlsx[tst].save(f'{tmp_path}/tmp_file.xlsx')
        to_test = openpyxl.load_workbook(filename=f'{tmp_path}/tmp_file.xlsx')
        # load excel file from data
        target = openpyxl.load_workbook(filename=f'tests/data/{tst}.xlsx')
        # Get sheet dimensions
        rows = range(target['Table'].min_row, target['Table'].max_row + 1)
        cols = range(target['Table'].min_column, target['Table'].max_column + 1)
        identical = True
        for row in rows:
            for col in cols:
                if not identical:
                    raise ValueError('Mismatch between test files and excel workbooks.')
                cell_address = f'{openpyxl.utils.get_column_interval(col, col)[0]}{row}'
                # compare values
                if target['Table'][cell_address].value is None:
                    if to_test['Table'][cell_address].value is None:
                        next
                    else:
                        identical = False
                        if not identical:
                            raise ValueError(
                                'Mismatch between test files and excel workbooks.'
                            )
                        next
                identical = identical & (
                    target['Table'][cell_address].value
                    == to_test['Table'][cell_address].value
                )
                if not identical:
                    raise ValueError('Mismatch between test files and excel workbooks.')
                # compare styles
                identical = identical & (
                    target['Table'][cell_address].style
                    == to_test['Table'][cell_address].style
                )
                if not identical:
                    raise ValueError('Mismatch between test files and excel workbooks.')
                # borders
                identical = identical & (
                    (
                        target['Table'][cell_address].border.left,
                        target['Table'][cell_address].border.right,
                        target['Table'][cell_address].border.top,
                        target['Table'][cell_address].border.bottom,
                    )
                    == (
                        to_test['Table'][cell_address].border.left,
                        to_test['Table'][cell_address].border.right,
                        to_test['Table'][cell_address].border.top,
                        to_test['Table'][cell_address].border.bottom,
                    )
                )
                if not identical:
                    raise ValueError('Mismatch between test files and excel workbooks.')

                identical = identical & (
                    (
                        target['Table'][cell_address].border.left,
                        target['Table'][cell_address].border.right,
                        target['Table'][cell_address].border.top,
                        target['Table'][cell_address].border.bottom,
                    )
                    == (
                        to_test['Table'][cell_address].border.left,
                        to_test['Table'][cell_address].border.right,
                        to_test['Table'][cell_address].border.top,
                        to_test['Table'][cell_address].border.bottom,
                    )
                )

                if not identical:
                    raise ValueError('Mismatch between test files and excel workbooks.')

                identical = identical & (
                    (
                        target['Table'][cell_address].font.bold,
                        target['Table'][cell_address].font.size,
                    )
                    == (
                        to_test['Table'][cell_address].font.bold,
                        to_test['Table'][cell_address].font.size,
                    )
                )
                if not identical:
                    raise ValueError('Mismatch between test files and excel workbooks.')
