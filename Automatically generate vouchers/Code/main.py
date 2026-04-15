import subprocess
import sys

# 按顺序执行三个脚本
# 任意一个失败（returncode != 0）立即停止，不继续执行后续脚本
scripts = [
    "bank.py",   # 第一步：登录网银，下载昨日流水，移动到目标文件夹
    "vouchers.py",    # 第二步：清洗流水数据，按模板批量生成凭证文件，追加汇总表
    "zfi024.py",     # 第三步：SAP批量导入凭证，回填凭证号到汇总表J列
]

for script in scripts:
    print(f"\n===== 开始运行：{script} =====")
    result = subprocess.run([sys.executable, script], check=False)
    if result.returncode != 0:
        print(f"❌ {script} 运行失败，中止后续流程")
        break
    print(f"✅ {script} 完成")

print("\n全部完成")
input("按回车键关闭...")  # 防止窗口一闪而过
