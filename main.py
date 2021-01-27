from selenium import webdriver
import time
import xlrd
from datetime import date,datetime

def main():
    wd = webdriver.Chrome()
    wd.implicitly_wait(10)
    wd.get('https://consoleprod.csgmall.com.cn/o2om/consignment-query-supplier/list')
    wd.set_window_position(20, 40)
    wd.set_window_size(1100, 700)
    elementuser = wd.find_element_by_id('username')
    elementuser.send_keys('CSGSUP000019')
    elementpass = wd.find_element_by_id('password')
    elementpass.send_keys('25bd8e13ef48c061@Jdcom')
    elementbutton = wd.find_element_by_tag_name('button')
    elementbutton.click()
    time.sleep(1)
    enter = wd.find_element_by_class_name('ant-modal-close-x')
    enter.click()
    time.sleep(10)
    elementinput = wd.find_elements_by_class_name('ant-input')[0]
    xl = xlrd.open_workbook('1.xlsx') #excel
    table = xl.sheets()[0]

    rows = table.nrows
    print('列数',rows)
    #取总列数
    n = 0
    for n in range(rows):
        time.sleep(2)
        danhao = table.cell(n, 0).value
        # excel 取单号完毕，准备就绪，可以冲了
        elementinput.clear()
        elementinput.send_keys(danhao)
        time.sleep(1)
        elementfind = wd.find_elements_by_tag_name('button')[3]
        elementfind.click()
        time.sleep(1)
        # elementidenter = wd.find_element_by_link_text('XS20210115103916')
        elementidenter = wd.find_elements_by_css_selector('.ant-table-row.ant-table-row-level-0>td:nth-child(4)>a')
        # print(elementidenter,elementidenter[0].text)
        elementidenter[0].click()
        time.sleep(1)
        elemencaidan = wd.find_elements_by_css_selector('.page-head-operator>button:nth-child(5)')
        elemencaidan[0].click()
        time.sleep(1)
        caidan = table.cell(n,4).value
        # elementcaidan = wd.find_elements_by_css_selector('.themed-input-inner>span:nth-last-child(3)')
        ele_1 = '.ant-tabs-tabpane.ant-tabs-tabpane-active .page-container.ant-layout-content .page-content-wrap .ant-col-10:nth-child(3) input'
        elementcaidan = wd.find_element_by_css_selector(ele_1)
        elementcaidan.send_keys(caidan)
        # print(caidan,elementcaidan.get_attribute('value'))
        time.sleep(1)
        elementwuliu = wd.find_element_by_css_selector('.anticon.anticon-search')
        elementwuliu.click()
        time.sleep(1)
        wuliu = table.cell(n,3).value
        # print(wuliu)
        ele_2 = '.ant-form.ant-form-horizontal .ant-col-12:nth-child(2) input'#物流查询
        elementwuliuchaxun_1 = wd.find_element_by_css_selector(ele_2)
        elementwuliuchaxun_1.clear()
        elementwuliuchaxun_1.send_keys(wuliu)
        time.sleep(1)
        elementwuliuchaxun_2 = wd.find_element_by_css_selector('.lov-modal-btn-container>.ant-btn.ant-btn-primary')
        elementwuliuchaxun_2.click()
        time.sleep(1)
        ele_3 = '.ant-radio-input'
        elementwuliuchaxun_3 = wd.find_element_by_css_selector(ele_3)
        elementwuliuchaxun_3.click()
        time.sleep(1)
        ele_4 = '.ant-modal-footer .ant-btn.ant-btn-primary'
        elementwuliuchaxun_4 = wd.find_element_by_css_selector(ele_4)
        elementwuliuchaxun_4.click()
        time.sleep(1)
        #wuliu
        time.sleep(1)
        commodity_name = table.cell(n, 1).value#选择商品
        commodity_shuliang = table.cell(n, 2).value#选择数量
        print('商品信息',commodity_name,commodity_shuliang)
        goodsInfo = getname(wd,commodity_name)
        ele_6 = '.ant-table-row.components-edit-table-index-hzero-edit-table.ant-table-row-level-0>td:nth-child(5) input'
        elementchoice = wd.find_element_by_css_selector(ele_6)
        elementchoiceIsSelect = elementchoice
        if type(elementchoice) == type([0]):
            elementchoiceIsSelect = elementchoice[goodsInfo['index']]

        if commodity_name == goodsInfo['name']:
            # for i in range(commodity_shuliang):
            elementchoiceIsSelect.send_keys(int(commodity_shuliang))
        else:
            pass
        time.sleep(1)
        # time.sleep(2000)
        #保存
        elementconfig = wd.find_element_by_css_selector('.ant-card>.ant-card-body:nth-child(1) button')
        elementconfig.click()
        time.sleep(3)
        elementclose_1 = wd.find_element_by_css_selector('.ant-tabs-tab-active.ant-tabs-tab .anticon.anticon-close')
        elementclose_1.click()
        time.sleep(3)
        elementclose_2 = wd.find_element_by_css_selector('.ant-tabs-tab .anticon.anticon-close')
        elementclose_2.click()
        time.sleep(3)
        n = n + 1
        print('完成')
    else:
        print('失败')

    time.sleep(60)
    driver.quit()

def getname(wd,commodity_name):
    # shuliang
    ele_5 = '.ant-table-row.components-edit-table-index-hzero-edit-table.ant-table-row-level-0>td:nth-child(3)'
    elementchoice = wd.find_element_by_css_selector(ele_5)
    # commodity_name = table.cell(n, 1).value  # 选择商品
    html_name = ''
    index = 0
    print('webelement---',elementchoice)
    if type(elementchoice) == type([0]):
        for x in range(0, len(elementchoice)):
            print('当前webelement', commodity_name, x, elementchoice[x])
            if commodity_name == elementchoice[x].text:
                # 商品名称 html_name
                # 当前索引 x
                index = x
                html_name = elementchoice[x].text
    else:
        if commodity_name == elementchoice.text:
            html_name = elementchoice.text
    time.sleep(2)
    goodsInfo = {
        'name':html_name,
        'index':index
    }
    print('XXX',commodity_name,html_name)
    return goodsInfo

if __name__=='__main__':
    main()







