import win32com.client, time
from pathlib import Path
import pandas as pd


# ===================================基础配置========================================
# region 基础配置
path = Path(r"C:\Users\DELL\Desktop\销售会计1")
total_path = r'C:\Users\Yuyu\Desktop\收款\汇总表.xlsx'
total = pd.read_excel(total_path, sheet_name='汇总数据')
# endregion
#===================================结束=========================================


# ==================================连接SAP=============================================
# region 连接SAP
# 连接SAP，拿到session 选择第一个账号第二个页面
session = win32com.client.GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(1)
# 最大化窗口
session.findById("wnd[0]").maximize()
# endregion
# ===================================结束===============================================


#自定义导航回ZFI024初始页面（开始前和每次处理完都要调一次，防止页面状态残留） 多次使用定义函数方便调用
def zfi024(session):
    # 输入事务码/nzfi024
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nzfi024"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()
    time.sleep(1)
    # 取消勾选测试运行复选框
    session.findById("wnd[0]/usr/chkP_CB1").selected = False


# ==================================进入zfi024页面导入凭证=================================
# region 进入zfi024页面导入凭证
#调用自定义函数先进一次初始页面
zfi024(session)
#创建一个列表存储数据 凭证结果
results = []
# endregion

# -----------------------------凭证提交操作-------------------------------------------------
#排序遍历文件夹里每个xls文件
for file in sorted(path.glob("*.xls")):
    # 打开文件选择弹窗
    session.findById("wnd[0]").sendVKey(4)
    time.sleep(0.3)
    # 弹窗里点确认
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(0.3)
    # 点进路径输入框 再按F4进入文件浏览
    session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[1]").sendVKey(4)
    time.sleep(0.3)
    # 填入文件夹路径  填入文件名
    session.findById("wnd[2]/usr/ctxtDY_PATH").text = str(file.parent)
    session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = file.name
    time.sleep(0.5)
    # 确认路径和文件名
    session.findById("wnd[2]/tbar[0]/btn[0]").press()
    time.sleep(0.3)
    # 确认选择对号
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(0.3)
    # 点执行按钮提交按钮
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    time.sleep(3)

# ------------------------------------凭证返回信息提取-----------------------------------------
    # 1. 首先，定位到目标网格对象
    grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    # 2. 获取第一条消息（消息在 "MESSAGE" 列，行索引为 0）
    row_index = 0
    column_name = "MESSAGE"
    # 3.凭证提交异常处理 zfi024凭证提交后会出现返回空值即提交失败
    try:
        #获取SAP返回对应单元格内容
        message_content = grid.GetCellValue(row_index, column_name)
        print(f"获取到的消息内容: {message_content}")
    except Exception as e:
        message_content = "凭证未提交成功"
        print(f"获取单元格值时出错: {e}")

# ----------------------------------------返回信息判断回传-----------------------------------------
    # message_content 情况1： 凭证过帐成功: BKPFF 180000139410302026 PRDCLNT800
    if "凭证过帐成功" in message_content:
        #提取9位凭证号
        result_text = message_content.split("BKPFF ")[1].split()[0][:9]
        # 数据追加到列表
        results.append(result_text)
    # message_content 情况2： 凭证错误: BKPFF $ PRDCLNT800
    elif "凭证错误" in message_content:
        # 直接返回凭证错误 客户冻结情况
        result_text = "凭证错误 客户冻结"
        results.append(result_text)
    else:
        # 没有获取到massage 返回为空 即凭证提交失败
        result_text = "凭证错误 未提交成功"
        results.append(result_text)
    print(f"完成：{file.name}")
    # 处理完回到初始页面，准备下一个文件
    zfi024(session)

# ===================================结束===============================================


# =================================凭证状态回填到凭证汇总表格==========================================
# region 凭证状态回填到凭证汇总表格
# 筛选出汇总表凭证号为空的列内容
empty_indices = total.index[total["凭证号"].isna() | (total["凭证号"] == "")]
# 将获取到的凭证返回信息传回数据表
for i, cert in zip(empty_indices, results):
    if cert:
        total.at[i,"凭证号"] = cert
with pd.ExcelWriter(total_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    total.to_excel(writer, sheet_name='汇总数据', index=False)
print("全部完成")
# endregion

# ===================================结束======================================================


