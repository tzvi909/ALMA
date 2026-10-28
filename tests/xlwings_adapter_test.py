"""Run the file-mode scenarios through XlwingsIO (the code path used inside
Excel) with tests/fake_xlwings.py standing in for xlwings, including main()."""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import fake_xlwings  # noqa: E402
sys.modules["xlwings"] = fake_xlwings
import xlwings_file_mode as fm  # noqa: E402

AX = fm.AX


def main(out):
    orig = AX.cli

    def cli_excel(argv):            # every step drives the "open Excel workbook" through XlwingsIO
        return orig(argv + ["--excel"])
    AX.cli = cli_excel
    bad = fm.main(out)
    book = fake_xlwings.Book(os.path.join(out, "ALMA_xlwings_test.xlsx"))
    # main(): the entry point of xlwings "Run main"
    book.set_mock_caller()
    a = AX.Alma(AX.XlwingsIO(book))
    a.set_setting("action", "Self-test")
    a.set_setting("quiet", False)
    msg = AX.main()
    print("main():", msg, "| alert shown:", book.app.alerts[-1:] if book.app.alerts else None)
    print("screen_updating restored:", book.app.screen_updating, "| status bar:", book.app.status_bar)
    return bad


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
