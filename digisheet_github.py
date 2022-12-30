from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import tkinter
import tkinter.ttk as ttk
from tkinter import messagebox
import calendar
import datetime

''' グローバル変数 '''
driver = None

def isAppropriateStaffId(BEFORE_WORD, AFTER_WORD):
    ''' 適切なスタッフ ID ( 7 ケタ以下 かつ 数字)しか入力させません。'''
    return ((AFTER_WORD.isdecimal()) and (len(AFTER_WORD)<=7)) or (len(AFTER_WORD) == 0)

def initChromeDriver():
    ''' ChromeDriver の初期設定を行います。 '''
    global driver
    PATH_CHROME_DRIVER = "./chromedriver"
    
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.use_chromium = True
    driver = webdriver.Chrome(executable_path=PATH_CHROME_DRIVER, options=options)

def loginToDigisheet(STAFF_CODE, PASSWORD):
    ''' digisheet にログインします。 '''
    global driver
    COMPANY_CODE = ""
    URL_DIGISHEET = ""

    driver.get(URL_DIGISHEET)

    # ログイン情報を入力し、ログインボタンを押下します。
    XPATH_COMPANY_CODE = "/html/body/form/center/table/tbody/tr/td/table/tbody/tr[3]/td/input"
    elemCompanyCode = driver.find_element(By.XPATH, XPATH_COMPANY_CODE)
    elemCompanyCode.send_keys(COMPANY_CODE)

    XPATH_STAFF_ID = "/html/body/form/center/table/tbody/tr/td/table/tbody/tr[4]/td/input"
    elemStaffId = driver.find_element(By.XPATH, XPATH_STAFF_ID)
    elemStaffId.send_keys(STAFF_CODE)

    XPATH_PASSWORD = "/html/body/form/center/table/tbody/tr/td/table/tbody/tr[5]/td/input"
    elemPassword = driver.find_element(By.XPATH, XPATH_PASSWORD)
    elemPassword.send_keys(PASSWORD)

    XPATH_LOGIN_BTN = "/html/body/form/center/table/tbody/tr/td/table/tbody/tr[6]/td/table/tbody/tr/th[1]/input"
    login_button = driver.find_element(By.XPATH, XPATH_LOGIN_BTN)
    login_button.click()

def clickWorkReport():
    ''' 勤務報告に遷移します。 '''
    global driver
    
    # 画面左側のメニューは frame で埋め込まれているので、まずは操作対象を frame に切り替える必要があります。
    ELEM_NAME_MENU_ON_LEFT = "menu"
    menuFrame = driver.find_element(By.NAME, ELEM_NAME_MENU_ON_LEFT)
    driver.switch_to.frame(menuFrame)

    # 「勤怠報告」をクリックします。
    XPATH_WORK_TIME_REPORT = "/html/body/form/div[4]/table/tbody/tr[7]/td/a"
    elemWorkTimeReport = driver.find_element(By.XPATH, XPATH_WORK_TIME_REPORT)
    elemWorkTimeReport.click()

    # 操作対象を初期値に戻します。戻さない場合、別の frame の操作に切り替えられませんでした。
    driver.switch_to.default_content()

def isRegistered(rowNumOfTargetDay):
    ''' 対象の日の勤務データが登録済みか確認します。 '''
    global driver

    xpathTdWorkingHourForEachDay = "/html/body/form/table/tbody/tr[7]/td/table/tbody/tr[" + rowNumOfTargetDay + "]/td[9]/font"
    tdWorkingHourForEachDay = driver.find_element(By.XPATH, xpathTdWorkingHourForEachDay)
    registeredWorkingHourOfTargetDay = tdWorkingHourForEachDay.get_attribute("textContent")

    # 空白に見える、勤務時間列の要素の文字数を len メソッドで確認したところ、3 が返ってきた。0ではなかった。
    # おそらく、「&nbsp;」とその前後の空白を計測しているのではないかと考えられる。
    # 従って、4 以上の場合は勤務情報が登録されていると判断ます。
    if len(registeredWorkingHourOfTargetDay) > 3 :
        return True
    else:
        return False

def isWorkDay(ROW_NUM_OF_TARGET_DAY):
    ''' 対象の日が営業日かどうかを確認します。 '''
    global driver

    WHITE = "rgba(255, 255, 255, 1)"
    
    xpathTrForTargetDay= "/html/body/form/table/tbody/tr[7]/td/table/tbody/tr[" + ROW_NUM_OF_TARGET_DAY + "]"
    rowOfTargetDay = driver.find_element(By.XPATH, xpathTrForTargetDay)
    bgcolorOfTrOfTargetDay = rowOfTargetDay.value_of_css_property("background-color")

    # 行の背景色が白い場合は営業日と判断します。 
    if (bgcolorOfTrOfTargetDay == WHITE):
        return True
    else:
        return False

def clickLinkToTargetDay(ROW_NUM_OF_TARGET_DAY):
    ''' 営業日のリンクをクリックします。 '''
    xpathAnchorTagOfTargetDay = "/html/body/form/table/tbody/tr[7]/td/table/tbody/tr[" + ROW_NUM_OF_TARGET_DAY + "]/td[3]/a"
    linkToTargetDay = driver.find_element(By.XPATH, xpathAnchorTagOfTargetDay)
    linkToTargetDay.click()

def registerWorkingHourOfTargetDay(START_HOUR, START_MINUTE, END_HOUR, END_MINUTE):
    ''' 勤務データを登録します。 '''
    global driver

    ELEM_NAME_START_HOUR_OF_WORK = "HourStart"
    ELEM_NAME_START_MINUTE_OF_WORK = "MinuteStart"
    ELEM_NAME_END_HOUR_OF_WORK = "HourEnd"
    ELEM_NAME_END_MINUTE_OF_WORK = "MinuteEnd"

    # 始業/終業時間のプルダウンを取得します。
    elemStartHour = Select(driver.find_element(By.NAME, ELEM_NAME_START_HOUR_OF_WORK))
    elemStartMinute = Select(driver.find_element(By.NAME, ELEM_NAME_START_MINUTE_OF_WORK))
    elemEndHour = Select(driver.find_element(By.NAME, ELEM_NAME_END_HOUR_OF_WORK))
    elemEndMinute = Select(driver.find_element(By.NAME, ELEM_NAME_END_MINUTE_OF_WORK))

    # プルダウンの値を定時の時、分に変更します。
    elemStartHour.select_by_value(START_HOUR)
    elemStartMinute.select_by_value(START_MINUTE)
    elemEndHour.select_by_value(END_HOUR)
    elemEndMinute.select_by_value(END_MINUTE)

    # 登録ボタンを押下します。
    XPATH_REGISTRATION_BUTTON = '//*[@id="StaffWorkSheet"]/table/tbody/tr[3]/td/table/tbody/tr[8]/td/table/tbody/tr/td/input[1]'
    registrationButton = driver.find_element(By.XPATH, XPATH_REGISTRATION_BUTTON)
    registrationButton.click()

def deleteWorkingHourOfTargetDay():
    ''' 登録した勤務データの削除を行います。（開発時しか使いません。） '''
    XPATH_DELETE_WORKING_HOUR_BTN = '//*[@id="StaffWorkSheet"]/table/tbody/tr[3]/td/table/tbody/tr[8]/td/table/tbody/tr/td/input[2]'
    deleteWorkingHourBtn = driver.find_element(By.XPATH, XPATH_DELETE_WORKING_HOUR_BTN)
    deleteWorkingHourBtn.click()

def registerWorkingHourIfItsTargetDay(START_HOUR, START_MINUTE, END_HOUR, END_MINUTE, REGISTRATION_START_DAY, DAYS_OF_MONTH):
    ''' 設定された日以降の各営業日に、設定された勤務時間のデータを登録します。 '''
    global driver

    # 勤務報告画面を操作するため、操作対象（frame）を切り替えます。
    ELEM_NAME_MAIN_MENU = "main"
    mainFrame = driver.find_element(By.NAME, ELEM_NAME_MAIN_MENU)
    driver.switch_to.frame(mainFrame)

    # 勤務実績が未入力の日に、定時の勤務データを作成していきます。
    # 繰り返しの回数は、（その月の総日数 - 処理開始対象日 + 1）で算出する。
    # 最後の +1 についてですが、繰り返しの回数を最低でも 1 にするための調整です。
    for i in range(DAYS_OF_MONTH - int(REGISTRATION_START_DAY) + 1):
        
        # 一日の行は tr[2] から始まっているので、rowNumOfTargetDay が最小でも 2 から始まるように調整します。
        rowNumOfTargetDay = str(i + int(REGISTRATION_START_DAY) + 1)

        # 勤務データが登録済みの場合は次の繰り返し処理に移ります。
        if (isRegistered(rowNumOfTargetDay)):
            continue
        
        # 勤務データが未登録かつ営業日の場合はデータ登録処理を実行します。
        if (isWorkDay(rowNumOfTargetDay)):
            clickLinkToTargetDay(rowNumOfTargetDay)
            registerWorkingHourOfTargetDay(START_HOUR, START_MINUTE, END_HOUR, END_MINUTE)

    driver.switch_to.default_content()
    
    SCRIPT = "alert('処理が完了しました。');"
    driver.execute_script(SCRIPT)

def main(STAFF_CODE, PASSWORD, START_HOUR, START_MINUTE, END_HOUR, END_MINUTE, REGISTRATION_START_DAY, DAYS_OF_MONTH):
    ''' 一連の処理を main メソッドにまとめています。 '''
    ret = messagebox.askokcancel(title='実行前確認', message='設定した値に基づき勤務データを登録します。\n誤って登録した場合は手動で更新/削除してください。', icon='question')
    
    if ret == True:
        initChromeDriver()
        loginToDigisheet(STAFF_CODE, PASSWORD)
        clickWorkReport()
        registerWorkingHourIfItsTargetDay(START_HOUR, START_MINUTE, END_HOUR, END_MINUTE, REGISTRATION_START_DAY, DAYS_OF_MONTH)
    else:
        return

def showTkWindow():
    ''' 入力画面を表示します。 '''
    TODAY = datetime.date.today()
    THIS_YEAR = TODAY.year
    THIS_MONTH = TODAY.month
    DAYS_OF_MONTH = calendar.monthrange(THIS_YEAR, THIS_MONTH)[1]

    def resetValuesToDefault():
        ''' テキストボックスとコンボボックスを初期状態に戻します。 '''
        resetFlag = messagebox.askokcancel(title="入力値初期化確認", message="テキストボックスとコンボボックスの値を初期値に戻します。", icon='warning')
        
        if resetFlag == True:
            entryStaffId.delete(0, tkinter.END)
            entryPassword.delete(0, tkinter.END)
            comboStartHour["values"] = hours
            comboStartHour.current(9)
            comboStartMinute.current(4)
            comboEndHour["values"] = hours
            comboEndHour.current(18)
            comboEndMinute.current(0)
            comboRegistrationStartDay.current(0)
        else:
            return

    def closeTkWindow():
        ''' tkinter の窓を閉じてもいいか確認します。 '''
        closeFlag = messagebox.askokcancel(title="終了確認", message="このアプリを閉じてもいいですか。", icon='question')
    
        if closeFlag == True:
            tkWindow.destroy()
        else:
            return

    # 今月の日数のコンボボックス用リスト
    days = []
    for i in range(1, DAYS_OF_MONTH + 1, 1):
        days.append(i)

    # 時のコンボボックス用リスト
    hours = []
    for i in range(0, 47, 1):
        hours.append(i)

    # 分のコンボボックス用リスト
    minutes = []
    for i in range(0, 60, 5):
        minutes.append(i)

    # 画面表示設定
    tkWindow = tkinter.Tk()
    tkWindow.geometry("300x350")
    tkWindow.title("勤務データ登録するやつ")

    # 処理対象年月を表示します。
    textTargetYearAndMonth = "登録処理対象月：" + str(THIS_YEAR) + "年" + str(THIS_MONTH) + "月"
    labelTargetYearAndMonth = tkinter.Label(text=textTargetYearAndMonth)
    labelTargetYearAndMonth.place(x=20, y=20)

    # スタッフIDの入力
    labelStaffId = tkinter.Label(text="スタッフID")
    labelStaffId.place(x=20, y=50)
    entryStaffId = tkinter.Entry(width=15)
    vcmd = (entryStaffId.register(isAppropriateStaffId), '%s', '%P')
    entryStaffId.configure(validate='key', vcmd=vcmd)
    entryStaffId.place(x=90, y=50)

    # パスワードの入力
    labelPassword = tkinter.Label(text="パスワード")
    labelPassword.place(x=20, y=80)
    entryPassword = tkinter.Entry(show='*', width=15)
    entryPassword.place(x=90, y=80)

    # 始業時間のラベル
    labelStartHour = tkinter.Label(text="始業時間")
    labelStartHour.place(x=20, y=130)

    def setApproproateEndHours(event):
        ''' 設定された始業時間（時）により、終業時間（時）のコンボボックスに適切な値を設定する。 '''
        endHours = []
        for i in range(int(comboStartHour.get()), 47):
            endHours.append(i)

        comboEndHour["values"] = endHours

    # 始業時間（時）のコンボボックス
    varStartHours = tkinter.StringVar()
    comboStartHour = ttk.Combobox(tkWindow, state="readonly", width=3, textvariable=varStartHours)
    comboStartHour["values"] = hours
    comboStartHour.place(x=90, y=130)
    comboStartHour.current(9)
    comboStartHour.bind('<<ComboboxSelected>>', setApproproateEndHours)

    # 始業時間の時間を区切るコロンを表示します。
    labelColonForStartTime = tkinter.Label(text=":")
    labelColonForStartTime.place(x=135, y=130)

    # 始業時間（分）のコンボボックス
    comboStartMinute = ttk.Combobox(tkWindow, state="readonly", width=3)
    comboStartMinute["values"] = minutes
    comboStartMinute.place(x=145, y=130)
    comboStartMinute.current(4)

    # 終業時間のラベル
    labelEndHour = tkinter.Label(text="終業時間")
    labelEndHour.place(x=20, y=160)

    def setApproproateStartHours(event):
        ''' 設定された終業時間（時）により、始業時間（時）のコンボボックスに適切な値を設定する。 '''
        startHours = []
        for i in range(0, int(comboEndHour.get()) + 1):
            startHours.append(i)

        comboStartHour["values"] = startHours

    # 終業時間（時）のコンボボックス
    varEndHours = tkinter.StringVar()
    comboEndHour = ttk.Combobox(tkWindow, state="readonly", width=3, textvariable=varEndHours)
    comboEndHour["values"] = hours
    comboEndHour.place(x=90, y=160)
    comboEndHour.current(18)
    comboEndHour.bind('<<ComboboxSelected>>', setApproproateStartHours)

    # 終業時間の時間を区切るコロンを表示します。
    labelColonForEndTime = tkinter.Label(text=":")
    labelColonForEndTime.place(x=135, y=160)

    # 終業時間（分）のコンボボックス
    comboEndMinute = ttk.Combobox(tkWindow, state="readonly", width=3)
    comboEndMinute["values"] = minutes
    comboEndMinute.place(x=145, y=160)
    comboEndMinute.current(0)

    # 登録開始日のラベル
    labelRegistrationStartDay = tkinter.Label(text="登録処理開始日")
    labelRegistrationStartDay.place(x=20, y=200)

    # 登録開始日のコンボボックス
    comboRegistrationStartDay = ttk.Combobox(tkWindow, state="readonly", width=3)
    comboRegistrationStartDay["values"] = days
    comboRegistrationStartDay.place(x=125, y=200)
    comboRegistrationStartDay.current(0)

    labelForRegistrationStartDay = tkinter.Label(text="日")
    labelForRegistrationStartDay.place(x=170, y=200)

    # 処理実行ボタン。
    button = tkinter.Button(tkWindow
                            , text = "勤務データ登録"
                            , command = lambda:main(
                                entryStaffId.get()
                                , entryPassword.get()
                                , comboStartHour.get()
                                , comboStartMinute.get()
                                , comboEndHour.get()
                                , comboEndMinute.get()
                                , comboRegistrationStartDay.get()
                                , DAYS_OF_MONTH
                            )
                            , foreground="#feb81c"
                            , background="#001f33"
                            , activeforeground="#001f33"
                            , activebackground="#feb81c")
    button.place(x=90, y=230)

    clearButton = tkinter.Button(tkWindow, text="入力値初期化", command=lambda:resetValuesToDefault())
    clearButton.place(x=50, y=265)

    closeWindowButton = tkinter.Button(tkWindow, text="閉じる", command=lambda:closeTkWindow())
    closeWindowButton.place(x=140, y=265)

    tkWindow.mainloop()

showTkWindow()

''' ゴミ置き場
# テーブルの行数から繰り返す回数を算出します。テーブル全体の行数からヘッダーとフッターの2行分を減算します。
HEADER_AND_FOOTER_ROWS = 2
XPATH_WORK_INFO_TABLE = "/html/body/form/table/tbody/tr[7]/td/table/tbody/tr"
numOfRowsInWorkInfoTable = len(driver.find_elements(By.XPATH, XPATH_WORK_INFO_TABLE)) - HEADER_AND_FOOTER_ROWS

# tkinter は、toolkit interfaceの略らしい。

# nbsp は no-breaking space の略らしい、覚えやすい。
'''