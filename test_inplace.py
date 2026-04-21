"""
Smoke test: verify that in-place UserForm modification works against the
existing SparePartsInventory_v2.xlsm without recreating forms.

Tries three operations that the real updater would need:
  1. Change the form's Height property
  2. Move an existing control (lstRes) to a new Top position
  3. Add a new Label control to the form

Leaves the file UNCHANGED (does not save).
"""

import os
import time
import xlwings as xw

import build_inventory as bi

PATH = r'c:\Users\avita\Claude_Projects\InventoryManager\SparePartsInventory_v2.xlsm'


def main():
    bi.enable_vba_access()
    os.system("taskkill /F /IM excel.exe >nul 2>&1")
    time.sleep(1)

    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False
    ok = {"height": False, "move": False, "add_ctrl": False}
    try:
        wb = app.books.open(PATH)
        print(f"Opened: {PATH}")

        vb = wb.api.VBProject
        comp = vb.VBComponents("frmSearch")
        d = comp.Designer
        print(f"  Got designer for frmSearch, control count={d.Controls.Count}")

        # Test 1: change form Height via Properties
        try:
            comp.Properties("Height").Value = 470
            print(f"  OK  [test 1] set frmSearch Height=470")
            ok["height"] = True
        except Exception as e:
            print(f"  FAIL[test 1] set Height: {e}")

        # Test 2: move an existing control (lstRes)
        try:
            for i in range(1, d.Controls.Count + 1):
                c = d.Controls(i)
                if c.Name == "lstRes":
                    orig_top = c.Top
                    c.Top = orig_top + 5
                    c.Top = orig_top  # restore
                    print(f"  OK  [test 2] moved lstRes (orig Top={orig_top})")
                    ok["move"] = True
                    break
            else:
                print("  FAIL[test 2] lstRes not found")
        except Exception as e:
            print(f"  FAIL[test 2] move control: {e}")

        # Test 3: add a new Label control to the form
        try:
            new_ctrl = d.Controls.Add("Forms.Label.1", "hdrTest", True)
            new_ctrl.Left = 6
            new_ctrl.Top = 48
            new_ctrl.Width = 70
            new_ctrl.Height = 14
            new_ctrl.Caption = "בדיקה"
            print(f"  OK  [test 3] added label hdrTest to frmSearch")
            ok["add_ctrl"] = True
            # Remove the test label so the form stays clean
            try:
                d.Controls.Remove("hdrTest")
                print("        (test label removed)")
            except Exception as e:
                print(f"        WARN: could not remove test label: {e}")
        except Exception as e:
            print(f"  FAIL[test 3] add control: {e}")

        # Do NOT save — close without saving
        wb.close()
    except Exception as e:
        import traceback
        print(f"ERROR: {e}")
        traceback.print_exc()
    finally:
        try:
            app.quit()
        except Exception:
            pass

    print("\n=== RESULT ===")
    for k, v in ok.items():
        print(f"  {k}: {'OK' if v else 'FAIL'}")
    if all(ok.values()):
        print("\nIn-place approach is viable.")
    else:
        print("\nIn-place approach will NOT work — at least one operation failed.")


if __name__ == "__main__":
    main()
