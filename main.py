from jsonSector import create_json_file, ReadSettingsExcel
from Models.Signal import Signal


def main():
    settings = ReadSettingsExcel()
    for setting in settings:
        print(vars(setting))
    print(len(settings))


    UserTree = [
        Signal("AI/Положение РК №1  (FSG1)", "algGVL.AI.POS_RK1.AI_ToHMI.PV", "%", "Положение РК №1  (FSG1)"),
        Signal("AI/Положение РК №2  (FSG2)", "algGVL.AI.POS_RK2.AI_ToHMI.PV", "%", "Положение РК №2  (FSG2)"),
        Signal("fdgdgfdgdfgfdgfd/Положение РК №2  (FSG2)", "algGVL.fdgdfdfgfd.POS_RK2.AI_ToHMI.PV", "%", "fffff РК №2  (FSG2)")
        # Добавьте нужное количество объектов Signal зд
    ]
    create_json_file(UserTree)



if __name__ == "__main__":
    main()
