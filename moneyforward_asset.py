from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
import chromedriver_binary
import pandas as pd
from getpass import getpass
import time

# web driverの初期設定
options = webdriver.ChromeOptions()
options.add_argument('--headless')
chrome_driver = webdriver.Chrome(options=options)


def login(driver=chrome_driver):
    """
    Money forward MEにログインする
    :param driver: webdriver
    :return: None
    """
    url = "https://moneyforward.com/users/sign_in"

    # 名前とパスワードの入力を受け付ける
    user_name = input('Your email address: ')
    password = getpass('Your password: ')

    driver.get(url)
    # メールアドレスを入力
    element_username = driver.find_element_by_id("sign_in_session_service_email")
    element_username.clear()
    element_username.send_keys(user_name)
    # 　パスワードを入力
    element_pass = driver.find_element_by_id("sign_in_session_service_password")
    element_pass.clear()
    element_pass.send_keys(password)
    # 　ログインボタンを押す
    frm = driver.find_element_by_name("commit")
    frm.click()

    # ログインに成功したか/失敗したかを表示
    time.sleep(5)
    if driver.current_url == 'https://moneyforward.com/session':
        print('ログイン失敗：メールアドレスもしくはパスワードが間違っています')
    else:
        print('ログイン成功')


def logout(driver=chrome_driver):
    '''
    Money forward MEからログアウトする
    :param driver: web driver
    :return: None
    '''
    ele_logout = driver.find_element_by_link_text('ログアウト')
    ele_logout.click()
    print('ログアウトしました')


def get_grouplist(driver=chrome_driver):
    """
    グループ一覧を取得する
    :param driver: web driver
    :return: list, list of group name
    """
    url_home = 'https://moneyforward.com/'
    driver.get(url_home)

    group_element = driver.find_element_by_name('group_id_hash')
    group_select_element = Select(group_element)

    group_list = []
    for ele in group_select_element.options:
        if ele.text != 'グループの追加・編集':
            group_list.append(ele.text)

    return group_list


def get_currentgroup(driver=chrome_driver):
    """
    現在選択中のグループを取得する
    :param driver:web driver
    :return: str, group name of current selected group
    """
    group_element = driver.find_element_by_name('group_id_hash')
    group_select_element = Select(group_element)

    selected_option = group_select_element.all_selected_options

    return selected_option[0].text


def select_group(group_name='グループ選択なし', driver=chrome_driver):
    '''
    Moneyforward MEのグループを選択する
    :param group_name: str, グループ名
    :param driver: web driver
    :return: None
    '''

    group_element = driver.find_element_by_name('group_id_hash')
    group_select_element = Select(group_element)

    try:
        group_select_element.select_by_visible_text(group_name)
        print('グループの指定ができました:', group_name)
    except:
        print('指定したグループ名が存在しません')


def f_remove_yen(x):
    '''
    文字列から円を取り除く(1,000円→1,000)
    :param x: str, 円がある文字列
    :return: int, 整数
    '''
    return int(str(x).replace('円', '').replace(',', ''))


def get_asset_summary(group_name='グループ選択なし', driver=chrome_driver):
    '''
    Moneyforward MEの資産サマリの表を取得して、DataFrame形式で出力する
    :param group_name:str, グループ名
    :param driver: web driver
    :return: dataframe,　資産のサマリの表
    '''
    # グループを選択する
    select_group(group_name, driver)

    url_bs = "https://moneyforward.com/bs/portfolio"
    driver.get(url_bs)
    # 資産のサマリデータを取得
    try:
        df_bssummary = pd.read_html(driver.page_source, attrs={'class': 'table table-bordered'})[0]
        df_bssummary.columns = ['種別', '金額', '比率']
        df_bssummary['評価額（円）'] = df_bssummary['金額'].map(f_remove_yen)
        df_bssummary = df_bssummary[['種別', '評価額（円）']]
        print('サマリの表データが見つかりました')
        return df_bssummary
    except ValueError:
        print('サマリの表データが見つかりませんでした')


def get_asset_rawtable(group_name='グループ選択なし', driver=chrome_driver):
    '''
    資産の表を取得する
    :param group_name:str, 表を取得する対象のグループ名
    :param driver: ドライバー
    :return: dictionary, keyが資産の分類名(例：'預金・現金・仮想通貨）, Valueが表(Dataframe形式）
    '''
    # グループの選択
    select_group(group_name, driver)

    # 資産のページを取得
    dict_assetdf ={}
    url_bs = "https://moneyforward.com/bs/portfolio"
    driver.get(url_bs)

    # 預金・現金・仮想通貨データを取得
    try:
        page_source_webelement = driver.find_element_by_id("portfolio_det_depo").get_attribute('innerHTML')
        df_det_depo = pd.read_html(page_source_webelement, attrs={'class': 'table table-bordered table-depo'})[0]
        dict_assetdf['預金・現金・仮想通貨'] = df_det_depo
        print('預金・現金・仮想通貨の表データが見つかりました')
    except (ValueError, NoSuchElementException):
        print('預金・現金・仮想通貨の表データが見つかりませんでした')

    # 株式（現物）データを取得
    try:
        page_source_webelement = driver.find_element_by_id("portfolio_det_eq").get_attribute('innerHTML')
        df_eq = pd.read_html(page_source_webelement, attrs={'class': 'table table-bordered table-eq'})[0]
        dict_assetdf['株式（現物）'] = df_eq
        print('株式（現物）の表データが見つかりました')
    except (ValueError, NoSuchElementException):
        print('株式（現物）の表データが見つかりませんでした')

    # 株式（信用）データを取得
    try:
        page_source_webelement = driver.find_element_by_id("portfolio_det_mgn").get_attribute('innerHTML')
        df_det_mgn = pd.read_html(page_source_webelement, attrs={'class': 'table table-bordered table-mgn'})[0]
        dict_assetdf['株式（信用）'] = df_det_mgn
        print('株式（信用）の表データが見つかりました')
    except (ValueError, NoSuchElementException):
        print('株式（信用）の表データが見つかりませんでした')

    # 投信信託データを取得
    try:
        page_source_webelement = driver.find_element_by_id("portfolio_det_mf").get_attribute('innerHTML')
        df_det_mf = pd.read_html(page_source_webelement, attrs={'class': 'table table-bordered table-mf'})[0]
        dict_assetdf['投資信託'] = df_det_mf
        print('投資信託の表データが見つかりました')
    except (ValueError, NoSuchElementException):
        print('投資信託の表データが見つかりませんでした')

    # 債券データを取得
    try:
        page_source_webelement = chrome_driver.find_element_by_id("portfolio_det_bd").get_attribute('innerHTML')
        df_det_bd = pd.read_html(page_source_webelement, attrs={'class': 'table table-bordered table-bd'})[0]
        dict_assetdf['債券'] = df_det_bd
        print('債券の表データが見つかりました')
    except (ValueError, NoSuchElementException):
        print('債券の表データが見つかりませんでした')

    # 先物・オプションデータを取得
    try:
        page_source_webelement = chrome_driver.find_element_by_id("portfolio_det_drv").get_attribute('innerHTML')
        df_det_drv = pd.read_html(page_source_webelement, attrs={'class': 'table table-bordered table-depo'})[0]
        dict_assetdf['先物・オプション'] = df_det_drv
        print('先物・オプションの表データが見つかりました')
    except (ValueError, NoSuchElementException):
        print('先物・オプションの表データが見つかりませんでした')

    # FXデータを取得
    try:
        page_source_webelement = chrome_driver.find_element_by_id("portfolio_det_fx").get_attribute('innerHTML')
        df_det_fx = pd.read_html(page_source_webelement, attrs={'class': 'table table-bordered table-depo'})[0]
        dict_assetdf['FX'] = df_det_fx
        print('FXの表データが見つかりました')
    except (ValueError, NoSuchElementException):
        print('FXの表データが見つかりませんでした')

    # 保険データを取得
    try:
        page_source_webelement = chrome_driver.find_element_by_id("portfolio_det_ins").get_attribute('innerHTML')
        df_det_ins = pd.read_html(page_source_webelement, attrs={'class': 'table table-bordered table-pns'})[0]
        dict_assetdf['保険'] = df_det_ins
        print('保険の表データが見つかりました')
    except (ValueError, NoSuchElementException):
        print('保険の表データが見つかりませんでした')

    # 不動産データを取得
    try:
        page_source_webelement = chrome_driver.find_element_by_id("portfolio_det_re").get_attribute('innerHTML')
        df_det_re = pd.read_html(page_source_webelement, attrs={'class': 'table table-bordered table-re'})[0]
        dict_assetdf['不動産'] = df_det_re
        print('不動産の表データが見つかりました')
    except (ValueError, NoSuchElementException):
        print('不動産の表データが見つかりませんでした')

    # 年金データを取得
    try:
        page_source_webelement = chrome_driver.find_element_by_id("portfolio_det_pns").get_attribute('innerHTML')
        df_det_pns = pd.read_html(page_source_webelement, attrs={'class': 'table table-bordered table-pns'})[0]
        dict_assetdf['年金'] = df_det_pns
        print('年金の表データが見つかりました')
    except (ValueError, NoSuchElementException):
        print('年金の表データが見つかりませんでした')

    # その他の資産データを取得
    try:
        page_source_webelement = chrome_driver.find_element_by_id("portfolio_det_oth").get_attribute('innerHTML')
        df_det_oth = pd.read_html(page_source_webelement, attrs={'class': 'table table-bordered table-pns'})[0]
        dict_assetdf['その他の資産'] = df_det_oth
        print('その他の資産の表データが見つかりました')
    except (ValueError, NoSuchElementException):
        print('その他の資産の表データが見つかりませんでした')

    # ポイントデータを取得
    try:
        page_source_webelement = chrome_driver.find_element_by_id("portfolio_det_po").get_attribute('innerHTML')
        df_det_po = pd.read_html(page_source_webelement, attrs={'class': 'table table-bordered table-pns'})[0]
        dict_assetdf['ポイント'] = df_det_po
        print('ポイントの表データが見つかりました')
    except (ValueError, NoSuchElementException):
        print('ポイントの表データが見つかりませんでした')

    return dict_assetdf


def get_asset_concatTable(group_name='グループ選択なし', driver=chrome_driver):
    '''
    各種類の資産のテーブルでの商品名と評価額を取得し、結合したテーブルにして出力
    :param group_name: str, テーブルを取得する対象のグループ名
    :param driver: web driver
    :return: Dataframe, 結合したテーブル
    '''

    # グループの選択、種別ごとの資産の表を取得
    dict_assetdf_rawtable = get_asset_rawtable(group_name, driver=chrome_driver)

    list_assetdf = []

    # 預金・現金・仮想通貨データの表を整形
    if '預金・現金・仮想通貨' in dict_assetdf_rawtable:
        df_depo = dict_assetdf_rawtable['預金・現金・仮想通貨']
        df_depo['評価額（円）'] = df_depo['残高'].map(f_remove_yen)
        df_depo = df_depo.rename(columns={'種類・名称': '名称'})
        df_depo['種別'] = '預金・現金・仮想通貨'
        df_depo = df_depo[['種別', '名称', '保有金融機関', '評価額（円）']]
        list_assetdf.append(df_depo)
        print('預金・現金・仮想通貨の表を整形しました')
    else:
        pass

    # 株式（現物）データを整形
    if '株式（現物）' in dict_assetdf_rawtable:
        df_eq = dict_assetdf_rawtable['株式（現物）']
        df_eq['評価額（円）'] = df_eq['評価額'].map(f_remove_yen)
        df_eq = df_eq.rename(columns={'銘柄名': '名称'})
        df_eq['種別'] = '株式（現物）'
        df_eq = df_eq[['種別', '名称', '保有金融機関', '評価額（円）']]
        list_assetdf.append(df_eq)
        print('株式（現物）の表を整形しました')
    else:
        pass

    # 株式（信用）データを整形
    if '株式（信用）' in dict_assetdf_rawtable:
        df_mgn = dict_assetdf_rawtable['株式（信用）']
        df_mgn['評価額（円）'] = df_mgn['現在値']
        df_mgn = df_mgn.rename(columns={'銘柄名': '名称'})
        df_mgn['種別'] = '株式（信用）'
        df_mgn = df_mgn[['種別', '名称', '保有金融機関', '評価額（円）']]
        list_assetdf.append(df_mgn)
        print('株式（信用）の表を整形しました')
    else:
        pass

    # 投信信託データを整形
    if '投資信託' in dict_assetdf_rawtable:
        df_mf = dict_assetdf_rawtable['投資信託']
        df_mf['評価額（円）'] = df_mf['評価額'].map(f_remove_yen)
        df_mf = df_mf.rename(columns={'銘柄名': '名称'})
        df_mf['種別'] = '投資信託'
        df_mf = df_mf[['種別', '名称', '保有金融機関', '評価額（円）']]
        list_assetdf.append(df_mf)
        print('投資信託の表を整形しました')
    else:
        pass

    # 債券データを整形
    if '債券' in dict_assetdf_rawtable:
        df_bd = dict_assetdf_rawtable['債券']
        df_bd['評価額（円）'] = df_bd['評価額'].map(f_remove_yen)
        df_bd = df_bd.rename(columns={'銘柄名': '名称'})
        df_bd['種別'] = '債券'
        df_bd = df_bd[['種別', '名称', '保有金融機関', '評価額（円）']]
        list_assetdf.append(df_bd)
        print('債券の表を整形しました')
    else:
        pass

    # 先物・オプションデータを整形
    if '先物・オプション' in dict_assetdf_rawtable:
        df_drv = dict_assetdf_rawtable['先物・オプション']
        df_drv['評価額（円）'] = df_drv['残高'].map(f_remove_yen)
        # df_drv = df_drv.rename(columns={'銘柄名': '名称'})
        df_drv['種別'] = '先物・オプション'
        df_drv = df_drv[['種別', '名称', '保有金融機関', '評価額（円）']]
        list_assetdf.append(df_drv)
        print('先物・オプションの表を整形しました')
    else:
        pass

    # FXデータを整形
    if 'FX' in dict_assetdf_rawtable:
        df_fx = dict_assetdf_rawtable['FX']
        df_fx['評価額（円）'] = df_fx['残高'].map(f_remove_yen)
        df_fx['種別'] = 'FX'
        df_fx = df_fx[['種別', '名称', '保有金融機関', '評価額（円）']]
        list_assetdf.append(df_fx)
        print('FXの表を整形しました')
    else:
        pass

    # 保険データを整形
    if '保険' in dict_assetdf_rawtable:
        df_ins = dict_assetdf_rawtable['保険']
        df_ins['評価額（円）'] = df_ins['現在価値'].map(f_remove_yen)
        df_ins['種別'] = '保険'
        df_ins['保有金融機関'] = 'なし'
        df_ins = df_ins[['種別', '名称', '保有金融機関', '評価額（円）']]
        list_assetdf.append(df_ins)
        print('保険の表を整形しました')
    else:
        pass

    # 不動産データを整形
    if '不動産' in dict_assetdf_rawtable:
        df_re = dict_assetdf_rawtable['不動産']
        df_re['評価額（円）'] = df_re['現在価値'].map(f_remove_yen)
        df_re['種別'] = '不動産'
        df_re['保有金融機関'] = 'なし'
        df_re = df_re[['種別', '名称', '保有金融機関', '評価額（円）']]
        list_assetdf.append(df_re)
        print('不動産の表を整形しました')
    else:
        pass

    # 年金データを整形
    if '年金' in dict_assetdf_rawtable:
        df_pns = dict_assetdf_rawtable['年金']
        df_pns['評価額（円）'] = df_pns['現在価値'].map(f_remove_yen)
        df_pns['種別'] = '年金'
        df_pns['保有金融機関'] = 'なし'
        df_pns = df_pns[['種別', '名称', '保有金融機関', '評価額（円）']]
        list_assetdf.append(df_pns)
        print('年金の表を整形しました')
    else:
        pass

    # その他の資産データを整形
    if 'その他の資産' in dict_assetdf_rawtable:
        df_oth = dict_assetdf_rawtable['その他の資産']
        df_oth['評価額（円）'] = df_oth['現在価値'].map(f_remove_yen)
        df_oth['種別'] = 'その他の資産'
        df_oth = df_oth[['種別', '名称', '保有金融機関', '評価額（円）']]
        list_assetdf.append(df_oth)
        print('その他の資産の表を整形しました')
    else:
        pass

    # ポイントデータを整形
    if 'ポイント' in dict_assetdf_rawtable:
        df_po = dict_assetdf_rawtable['ポイント']
        df_po['評価額（円）'] = df_po['現在の価値'].map(f_remove_yen)
        df_po['種別'] = 'ポイント'
        df_po = df_po[['種別', '名称', '保有金融機関', '評価額（円）']]
        list_assetdf.append(df_po)
        print('ポイントの表を整形しました')
    else:
        pass

    return pd.concat(list_assetdf).reset_index(drop=True)


if __name__ == '__main__':
    # 関数のテストプログラムを実行する
    login()
    time.sleep(1)

    grouplist = get_grouplist()
    print(grouplist)
    time.sleep(1)

    df_asset_summary = get_asset_summary(group_name=grouplist[0])
    print(df_asset_summary)
    time.sleep(1)

    dict_df_asset = get_asset_rawtable(group_name=grouplist[0])
    print(dict_df_asset)

    df_asset_detail = get_asset_concatTable(group_name=grouplist[0])
    print(df_asset_detail)
