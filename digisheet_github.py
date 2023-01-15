from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import tkinter
import tkinter.ttk as ttk
from tkinter import messagebox
import calendar
import datetime
import locale

''' グローバル変数 '''
# 処理の中で driver の設定を行った場合、処理の終了と同時に driver が終了した。それを回避するためにグローバル変数として扱っている。
driver = None

def initChromeDriver():
  ''' ChromeDriver の初期設定を行います。 '''
  global driver
  PATH_CHROME_DRIVER = "./chromedriver"
  
  options = webdriver.ChromeOptions()
  options.add_experimental_option('excludeSwitches', ['enable-logging'])
  options.use_chromium = True
  driver = webdriver.Chrome(executable_path=PATH_CHROME_DRIVER, options=options)

def logInToDigisheet(STAFF_CODE, PASSWORD):
  ''' digisheet にログインします。 '''
  global driver

  URL_DIGISHEET = ""
  driver.get(URL_DIGISHEET)

  XPATH_COMPANY_CODE = "/html/body/form/center/table/tbody/tr/td/table/tbody/tr[3]/td/input"
  elemCompanyCode = driver.find_element(By.XPATH, XPATH_COMPANY_CODE)
  COMPANY_CODE = ""
  elemCompanyCode.send_keys(COMPANY_CODE)

  XPATH_STAFF_ID = "/html/body/form/center/table/tbody/tr/td/table/tbody/tr[4]/td/input"
  elemStaffId = driver.find_element(By.XPATH, XPATH_STAFF_ID)
  elemStaffId.send_keys(STAFF_CODE)

  XPATH_PASSWORD = "/html/body/form/center/table/tbody/tr/td/table/tbody/tr[5]/td/input"
  elemPassword = driver.find_element(By.XPATH, XPATH_PASSWORD)
  elemPassword.send_keys(PASSWORD)

  XPATH_LOGIN_BUTTON = "/html/body/form/center/table/tbody/tr/td/table/tbody/tr[6]/td/table/tbody/tr/th[1]/input"
  login_button = driver.find_element(By.XPATH, XPATH_LOGIN_BUTTON)
  login_button.click()

def clickWorkReport():
  ''' 勤務報告に遷移します。 '''
  global driver
  
  # 画面左側のメニューは frame で埋め込まれているので、まずは操作対象を frame に切り替える必要があります。
  ELEM_NAME_MENU_ON_LEFT = "menu"
  menuFrame = driver.find_element(By.NAME, ELEM_NAME_MENU_ON_LEFT)
  driver.switch_to.frame(menuFrame)

  XPATH_WORK_TIME_REPORT = "/html/body/form/div[4]/table/tbody/tr[7]/td/a"
  elemWorkTimeReport = driver.find_element(By.XPATH, XPATH_WORK_TIME_REPORT)
  elemWorkTimeReport.click()

  # 操作対象を初期値に戻します。戻さない場合、別の frame の操作に切り替えられませんでした。
  driver.switch_to.default_content()

def isSpecialLeaveOrPaidVacation(ROW_NUM_OF_TARGET_DAY):
  ''' 有給休暇や特別休暇などを届け出ているか確認します。 A勤務でも空白でもなければ、何らかの申請をしていると判断します。 '''
  global driver

  WORK_PATTERN_A = "Ａ勤務"
  WORK_PATTERN_BLANK = " "

  xpathApplicationStatusForEachDay = "/html/body/form/table/tbody/tr[7]/td/table/tbody/tr[" + ROW_NUM_OF_TARGET_DAY + "]/td[8]/font"
  elemApplicationStatusOfTargetDay = driver.find_element(By.XPATH, xpathApplicationStatusForEachDay)
  applicationStatusOfTargetDay = elemApplicationStatusOfTargetDay.text

  if (applicationStatusOfTargetDay == WORK_PATTERN_A or applicationStatusOfTargetDay == WORK_PATTERN_BLANK):
    return False
  else:
    return True

def isRegistered(ROW_NUM_OF_TARGET_DAY):
  ''' 勤務データが登録済みか確認します。 '''
  global driver

  xpathTdWorkingHourForEachDay = "/html/body/form/table/tbody/tr[7]/td/table/tbody/tr[" + ROW_NUM_OF_TARGET_DAY + "]/td[9]/font"
  tdWorkingHourForEachDay = driver.find_element(By.XPATH, xpathTdWorkingHourForEachDay)
  registeredWorkingHourOfTargetDay = tdWorkingHourForEachDay.get_attribute("textContent")

  # 空白に見える、勤務時間列の要素の文字数を len メソッドで確認したところ、3 が返ってきた。0ではなかった。
  # おそらく、「&nbsp;」とその前後の空白を計測しているのではないかと考えられる。
  # 従って、4 以上の場合は勤務情報が登録されていると判断します。
  NUM_OF_CHARS_WHEN_BLANK = 3
  if len(registeredWorkingHourOfTargetDay) > NUM_OF_CHARS_WHEN_BLANK:
    return True
  else:
    return False

def isWorkDay(ROW_NUM_OF_TARGET_DAY):
  ''' 行の背景色が白い場合は営業日と判断します。 '''
  global driver

  xpathTrForEachDay = "/html/body/form/table/tbody/tr[7]/td/table/tbody/tr[" + ROW_NUM_OF_TARGET_DAY + "]"
  rowOfTargetDay = driver.find_element(By.XPATH, xpathTrForEachDay)
  bgcolorOfTrOfTargetDay = rowOfTargetDay.value_of_css_property("background-color")
  
  WHITE = "rgba(255, 255, 255, 1)"
  if (bgcolorOfTrOfTargetDay == WHITE):
    return True
  else:
    return False

def clickLinkToTargetDay(ROW_NUM_OF_TARGET_DAY):
  ''' 営業日のリンクをクリックします。 '''
  xpathAnchorTagOfTargetDay = "/html/body/form/table/tbody/tr[7]/td/table/tbody/tr[" + ROW_NUM_OF_TARGET_DAY + "]/td[3]/a"
  linkToTargetDay = driver.find_element(By.XPATH, xpathAnchorTagOfTargetDay)
  linkToTargetDay.click()

def applyForTelecommuting():
  ''' 在宅勤務の設定を行います。 '''

  XPATH_TELECOMMUTING = "//*[@id='StaffWorkSheet']/table/tbody/tr[3]/td/table/tbody/tr[5]/td/table/tbody/tr/td/select"
  elemTelecommuting = Select(driver.find_element(By.XPATH, XPATH_TELECOMMUTING))
  COMBO_VALUE_TELECOMMUTING = "0000000600"
  elemTelecommuting.select_by_value(COMBO_VALUE_TELECOMMUTING)

def deleteWorkingHourOfTargetDay():
  ''' 登録した勤務データの削除を行います。（開発用、登録したデータを消すのが面倒だから。） '''
  XPATH_DELETE_WORKING_HOUR_BUTTON = '//*[@id="StaffWorkSheet"]/table/tbody/tr[3]/td/table/tbody/tr[8]/td/table/tbody/tr/td/input[2]'
  deleteWorkingHourButton = driver.find_element(By.XPATH, XPATH_DELETE_WORKING_HOUR_BUTTON)
  deleteWorkingHourButton.click()

def registerWorkingHourOfTargetDay(START_HOUR, START_MINUTE, END_HOUR, END_MINUTE, TELECOMMUTING_FLAG):
  ''' 勤務データを登録します。 '''
  global driver

  ELEM_NAME_START_HOUR_OF_WORK = "HourStart"
  ELEM_NAME_START_MINUTE_OF_WORK = "MinuteStart"
  ELEM_NAME_END_HOUR_OF_WORK = "HourEnd"
  ELEM_NAME_END_MINUTE_OF_WORK = "MinuteEnd"

  elemStartHour = Select(driver.find_element(By.NAME, ELEM_NAME_START_HOUR_OF_WORK))
  elemStartMinute = Select(driver.find_element(By.NAME, ELEM_NAME_START_MINUTE_OF_WORK))
  elemEndHour = Select(driver.find_element(By.NAME, ELEM_NAME_END_HOUR_OF_WORK))
  elemEndMinute = Select(driver.find_element(By.NAME, ELEM_NAME_END_MINUTE_OF_WORK))

  elemStartHour.select_by_value(START_HOUR)
  elemStartMinute.select_by_value(START_MINUTE)
  elemEndHour.select_by_value(END_HOUR)
  elemEndMinute.select_by_value(END_MINUTE)

  if TELECOMMUTING_FLAG is True:
    applyForTelecommuting()
  
  XPATH_REGISTRATION_BUTTON = '//*[@id="StaffWorkSheet"]/table/tbody/tr[3]/td/table/tbody/tr[8]/td/table/tbody/tr/td/input[1]'
  registrationButton = driver.find_element(By.XPATH, XPATH_REGISTRATION_BUTTON)
  registrationButton.click()

def registerWorkingHourIfItsTargetDay(START_HOUR, START_MINUTE, END_HOUR, END_MINUTE, TELECOMMUTING_FLAG, REGISTRATION_START_DAY, REGISTRATION_END_DAY):
  ''' 設定された処理期間の各営業日に、設定された勤務時間のデータを登録します。 '''
  global driver

  # 勤務報告画面を操作するため、操作対象（frame）を切り替えます。
  ELEM_NAME_MAIN_MENU = "main"
  mainFrame = driver.find_element(By.NAME, ELEM_NAME_MAIN_MENU)
  driver.switch_to.frame(mainFrame)

  # 勤務実績が未入力の日に、勤務データを登録していきます。
  # 繰り返しの回数は、（処理終了日 - 処理開始日 + 1）で算出する。
  # 開始日と終了日が同じ日だとしても、繰り返しの回数を 1 にするために +1 で調整しています。
  for i in range(int(REGISTRATION_END_DAY) - int(REGISTRATION_START_DAY) + 1):
    
    # 一日（ついたち）の行は tr[2] から始まっているので、rowNumOfTargetDay が最小でも 2 から始まるように + 1 で調整します。
    rowNumOfTargetDay = str(i + int(REGISTRATION_START_DAY) + 1)

    if (isSpecialLeaveOrPaidVacation(rowNumOfTargetDay)):
      continue

    if (isRegistered(rowNumOfTargetDay)):
      continue

    if (isWorkDay(rowNumOfTargetDay)):
      clickLinkToTargetDay(rowNumOfTargetDay)
      registerWorkingHourOfTargetDay(START_HOUR, START_MINUTE, END_HOUR, END_MINUTE, TELECOMMUTING_FLAG)

  driver.switch_to.default_content()
  
  SCRIPT = "alert('処理が完了しました。');"
  driver.execute_script(SCRIPT)

def main(STAFF_CODE, PASSWORD, START_HOUR, START_MINUTE, END_HOUR, END_MINUTE, TELECOMMUTING_FLAG, REGISTRATION_START_DAY, REGISTRATION_END_DAY):
  ''' 一連の処理を main メソッドにまとめています。 '''
  ret = messagebox.askokcancel(title = '実行前確認', message = '設定した値に基づき勤務データを登録します。\n誤って登録した場合は手動で更新/削除してください。', icon = 'question')

  if ret is True:
    initChromeDriver()
    logInToDigisheet(STAFF_CODE, PASSWORD)
    clickWorkReport()
    registerWorkingHourIfItsTargetDay(START_HOUR, START_MINUTE, END_HOUR, END_MINUTE, TELECOMMUTING_FLAG, REGISTRATION_START_DAY, REGISTRATION_END_DAY)
  else:
    return

def showTkWindow():
  ''' 入力画面を表示します。 '''
  TODAY = datetime.date.today()
  THIS_YEAR = TODAY.year
  THIS_MONTH = TODAY.month
  DAYS_OF_MONTH = calendar.monthrange(THIS_YEAR, THIS_MONTH)[1]

  tkWindow = tkinter.Tk()
  tkWindow.geometry("300x400")
  tkWindow.title("勤務データ登録するやつ")

  targetYearAndMonth = "登録処理対象月：" + str(THIS_YEAR) + "年" + str(THIS_MONTH) + "月"
  labelTargetYearAndMonth = tkinter.Label(text = targetYearAndMonth)
  labelTargetYearAndMonth.place(x = 20, y = 20)

  def isAppropriateStaffId(BEFORE_WORD, AFTER_WORD):
    ''' 適切なスタッフ ID ( 7 ケタ以下 かつ 数字)しか入力させません。'''
    CHARACTOR_LIMIT = 7
    CHARACTOR_LOWER_LIMIT = 0

    return ((AFTER_WORD.isdecimal()) and (len(AFTER_WORD) <= CHARACTOR_LIMIT)) or (len(AFTER_WORD) == CHARACTOR_LOWER_LIMIT)

  labelStaffId = tkinter.Label(text = "スタッフID")
  labelStaffId.place(x = 20, y = 50)
  entryStaffId = tkinter.Entry(width = 15)
  vcmd = (entryStaffId.register(isAppropriateStaffId), '%s', '%P')
  entryStaffId.configure(validate = 'key', vcmd = vcmd)
  entryStaffId.place(x = 90, y = 50)

  labelPassword = tkinter.Label(text = "パスワード")
  labelPassword.place(x = 20, y = 80)
  entryPassword = tkinter.Entry(show = '*', width = 15)
  entryPassword.place(x = 90, y = 80)

  labelStartHour = tkinter.Label(text = "始業時間")
  labelStartHour.place(x = 20, y = 130)

  def setApproproateEndHours(event):
    ''' 設定された始業時間（時）により、終業時間（時）のコンボボックスに適切な値を設定する。 '''
    endHours = []
    for i in range(int(comboStartHour.get()), 47):
      endHours.append(i)

    comboEndHour["values"] = endHours

  def ProhibitsSettingEndTimeBeforeStartTime(event):
    ''' 終業時間を始業時間以前に設定することを禁じたい！ '''
    startHour = int(comboStartHour.get())
    endHour = int(comboEndHour.get())
    
    if startHour >= endHour:
      startMinute = int(comboStartMinute.get())
      endMinute = int(comboEndMinute.get())

      if startMinute >= endMinute:
        messagebox.showerror("警告", "終業時間が、始業時間以前になっているので直してください。")

  hours = []
  for i in range(0, 47, 1):
    hours.append(i)

  varStartHours = tkinter.StringVar()
  comboStartHour = ttk.Combobox(tkWindow, state = "readonly", width = 3, textvariable = varStartHours)
  comboStartHour["values"] = hours
  comboStartHour.place(x = 90, y = 130)
  comboStartHour.current(9)
  comboStartHour.bind('<<ComboboxSelected>>', setApproproateEndHours)
  comboStartHour.bind('<<ComboboxSelected>>', ProhibitsSettingEndTimeBeforeStartTime, "+")

  labelColonForStartTime = tkinter.Label(text = ":")
  labelColonForStartTime.place(x = 135, y = 130)

  minutes = []
  for i in range(0, 60, 5):
    minutes.append(i)

  comboStartMinute = ttk.Combobox(tkWindow, state = "readonly", width = 3)
  comboStartMinute["values"] = minutes
  comboStartMinute.place(x = 145, y = 130)
  # 20分を初期値にしたいので、4を設定している。分のコンボボックスは5分刻みになっている。
  comboStartMinute.current(4)
  comboStartMinute.bind('<<ComboboxSelected>>', ProhibitsSettingEndTimeBeforeStartTime)

  labelEndHour = tkinter.Label(text = "終業時間")
  labelEndHour.place(x = 20, y = 160)

  def setApproproateStartHours(event):
    ''' 設定された終業時間（時）により、始業時間（時）のコンボボックスに適切な値を設定する。 '''
    startHours = []
    for i in range(0, int(comboEndHour.get()) + 1):
      startHours.append(i)

    comboStartHour["values"] = startHours

  varEndHours = tkinter.StringVar()
  comboEndHour = ttk.Combobox(tkWindow, state = "readonly", width = 3, textvariable = varEndHours)
  comboEndHour["values"] = hours
  comboEndHour.place(x = 90, y = 160)
  comboEndHour.current(18)
  comboEndHour.bind('<<ComboboxSelected>>', setApproproateStartHours)
  comboEndHour.bind('<<ComboboxSelected>>', ProhibitsSettingEndTimeBeforeStartTime, "+")

  labelColonForEndTime = tkinter.Label(text = ":")
  labelColonForEndTime.place(x = 135, y = 160)

  comboEndMinute = ttk.Combobox(tkWindow, state = "readonly", width = 3)
  comboEndMinute["values"] = minutes
  comboEndMinute.place(x = 145, y = 160)
  comboEndMinute.current(0)
  comboEndMinute.bind('<<ComboboxSelected>>', ProhibitsSettingEndTimeBeforeStartTime)

  telecommutingFlag = tkinter.BooleanVar()
  checkTelecommuting = tkinter.Checkbutton(tkWindow, text = '在宅勤務', variable = telecommutingFlag)
  checkTelecommuting.place(x = 190, y = 160)

  labelRegistrationPeriod = tkinter.Label(text = "処理期間")
  labelRegistrationPeriod.place(x = 20, y = 200)

  def changeOptionsForRegistrationEndDay(event):
    ''' 設定された登録処理開始日により、登録処理終了日のコンボボックスに適切な値を設定する。 '''
    registrationEndDays = []
    for i in range(int(comboRegistrationStartDay.get()), DAYS_OF_MONTH + 1):
      registrationEndDays.append(i)

    comboRegistrationEndDay["values"] = registrationEndDays
  
  def getDayOfWeekInJapanese(YEAR, MONTH, DAY):
    ''' 日本語で曜日を取得します。 '''
    locale.setlocale(locale.LC_TIME, '')

    DAY_OF_WEEK = datetime.date(YEAR, MONTH, DAY)
    return "(" + DAY_OF_WEEK.strftime('%a') + ")"

  def setDayOfWeekForRegistrationStartDay(event):
    ''' 選択された日に応じた曜日を、処理開始日に設定する。 '''
    SELECTED_DAY = int(comboRegistrationStartDay.get())

    dayOfWeekForSelectedDay = getDayOfWeekInJapanese(THIS_YEAR, THIS_MONTH, SELECTED_DAY)
    labelDayOfWeekForRegistrationStartDay.configure(text = dayOfWeekForSelectedDay)

  days = []
  for i in range(1, DAYS_OF_MONTH + 1, 1):
    days.append(i)

  varRegistrationStartDay = tkinter.StringVar()
  comboRegistrationStartDay = ttk.Combobox(tkWindow, state = "readonly", width = 3, textvariable = varRegistrationStartDay)
  comboRegistrationStartDay["values"] = days
  comboRegistrationStartDay.place(x = 90, y = 200)
  comboRegistrationStartDay.current(0)
  comboRegistrationStartDay.bind('<<ComboboxSelected>>', changeOptionsForRegistrationEndDay)
  comboRegistrationStartDay.bind('<<ComboboxSelected>>', setDayOfWeekForRegistrationStartDay, "+")

  dayOfWeekForRegistrationStartDay = getDayOfWeekInJapanese(THIS_YEAR, THIS_MONTH, 1)
  labelDayOfWeekForRegistrationStartDay = tkinter.Label(text = dayOfWeekForRegistrationStartDay)
  labelDayOfWeekForRegistrationStartDay.place(x = 135, y = 200)

  labelWaveDashForRegistrationPeriod = tkinter.Label(text = "～")
  labelWaveDashForRegistrationPeriod.place(x = 165, y = 200)

  def changeOptionsForRegistrationStartDay(event):
    ''' 設定された登録処理終了日により、登録処理開始日のコンボボックスに適切な値を設定する。 '''
    registrationStartDays = []
    for i in range(1, int(comboRegistrationEndDay.get()) + 1):
      registrationStartDays.append(i)

    comboRegistrationStartDay["values"] = registrationStartDays
  
  def setDayOfWeekForRegistrationEndDay(event):
    ''' 選択された日に応じた曜日を、処理終了日に設定する。 '''
    SELECTED_DAY = int(comboRegistrationEndDay.get())

    dayOfWeekForSelectedDay = getDayOfWeekInJapanese(THIS_YEAR, THIS_MONTH, SELECTED_DAY)
    labelDayOfWeekForRegistrationEndDay.configure(text = dayOfWeekForSelectedDay)
  
  varRegistrationEndDay = tkinter.StringVar()
  comboRegistrationEndDay = ttk.Combobox(tkWindow, state = "readonly", width = 3, textvariable = varRegistrationEndDay)
  comboRegistrationEndDay["values"] = days
  comboRegistrationEndDay.place(x = 185, y = 200)
  comboRegistrationEndDay.current(DAYS_OF_MONTH - 1)
  comboRegistrationEndDay.bind('<<ComboboxSelected>>', changeOptionsForRegistrationStartDay)
  comboRegistrationEndDay.bind('<<ComboboxSelected>>', setDayOfWeekForRegistrationEndDay, "+")

  dayOfWeekForRegistrationEndDay = getDayOfWeekInJapanese(THIS_YEAR, THIS_MONTH, DAYS_OF_MONTH)
  labelDayOfWeekForRegistrationEndDay = tkinter.Label(text = dayOfWeekForRegistrationEndDay)
  labelDayOfWeekForRegistrationEndDay.place(x = 230, y = 200)

  # ボタンの色は 自社 のロゴのカラーコードを使用している。
  KAISYA_NAVY = "#001f33"
  KAISYA_ORANGE = "#feb81c"
  runStartButton = tkinter.Button(tkWindow
                          , text = "勤務データ登録"
                          , width = 30
                          , height = 2
                          , command = lambda:main(
                            entryStaffId.get()
                            , entryPassword.get()
                            , comboStartHour.get()
                            , comboStartMinute.get()
                            , comboEndHour.get()
                            , comboEndMinute.get()
                            , telecommutingFlag.get()
                            , comboRegistrationStartDay.get()
                            , comboRegistrationEndDay.get()
                          )
                          , foreground = KAISYA_ORANGE
                          , background = KAISYA_NAVY
                          , activeforeground = KAISYA_NAVY
                          , activebackground = KAISYA_ORANGE)
  runStartButton.place(x = 20, y = 230)

  def resetValuesToDefault():
    ''' 入力値を初期値に戻します。 '''
    resetFlag = messagebox.askokcancel(title = "入力値初期化確認", message = "入力値を初期値に戻します。", icon = 'warning')
    
    if resetFlag is True:
      entryStaffId.delete(0, tkinter.END)

      entryPassword.delete(0, tkinter.END)
      
      comboStartHour["values"] = hours
      comboStartHour.current(9)
      
      comboStartMinute.current(4)
      comboEndHour["values"] = hours
      comboEndHour.current(18)
      
      comboEndMinute.current(0)
      
      telecommutingFlag.set(False)
      
      comboRegistrationStartDay.current(0)
      
      dayOfWeekForRegistrationStartDay = getDayOfWeekInJapanese(THIS_YEAR, THIS_MONTH, 1)
      labelDayOfWeekForRegistrationStartDay.configure(text = dayOfWeekForRegistrationStartDay)

      comboRegistrationEndDay["values"] = days
      comboRegistrationEndDay.current(DAYS_OF_MONTH - 1)
            
      dayOfWeekForRegistrationEndDay = getDayOfWeekInJapanese(THIS_YEAR, THIS_MONTH, DAYS_OF_MONTH)
      labelDayOfWeekForRegistrationEndDay.configure(text = dayOfWeekForRegistrationEndDay)
    else:
      return

  clearButton = tkinter.Button(tkWindow, text = "入力値初期化", command = lambda:resetValuesToDefault())
  clearButton.place(x = 20, y = 280)

  def closeTkWindow():
    ''' tkinter の窓を閉じてもいいか確認します。 '''
    closeFlag = messagebox.askokcancel(title = "終了確認", message = "このアプリを閉じてもいいですか。", icon = 'question')
  
    if closeFlag is True:
      tkWindow.destroy()
    else:
      return
  
  closeWindowButton = tkinter.Button(tkWindow, text = "閉じる", command = lambda:closeTkWindow())
  closeWindowButton.place(x = 110, y = 280)

  tkWindow.mainloop()

showTkWindow()
