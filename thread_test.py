import sys
import os
import json
from PyQt5.QtCore import QCoreApplication
from main import FormatThread

def test_thread():
    """测试FormatThread类是否能正确处理txt文件"""
    print("开始测试FormatThread...")
    
    # 创建一个简单的测试配置
    config = {
        "margins": {"top": 2.5, "bottom": 2.5, "left": 3.0, "right": 3.0},
        "spacing": {"line_spacing": 1.5},
        "title_font": {"name": "黑体", "size": 16},
        "body_font": {"name": "宋体", "digit_font": "Times New Roman", "size": 10.5},
        "heading_levels": [
            {"font": "黑体", "size": 16, "bold": True},
            {"font": "黑体", "size": 14, "bold": True},
            {"font": "黑体", "size": 12, "bold": True},
            {"font": "黑体", "size": 10, "bold": True}
        ],
        "page_number": {
            "start": 1,
            "format": "PAGE"
        }
    }
    
    # 保存配置到临时文件
    with open('thread_test_config.json', 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=4)
    
    print("配置文件已创建")
    
    # 测试文件路径
    input_files = ['example.txt']
    output_dir = '.'  # 当前目录
    
    print(f"输入文件: {input_files}")
    print(f"输出目录: {output_dir}")
    
    # 创建并运行FormatThread
    app = QCoreApplication(sys.argv)  # 创建QApplication实例
    thread = FormatThread(input_files, output_dir, 'thread_test_config.json')
    
    # 连接信号
    def on_completed(output_dir, generated_files):
        print(f"处理完成，输出目录: {output_dir}")
        print(f"生成的文件: {generated_files}")
        app.quit()  # 退出应用程序
        
    def on_error(error_msg):
        print(f"处理出错: {error_msg}")
        # 添加更多调试信息
        import traceback
        traceback.print_exc()
        app.quit()  # 退出应用程序
        
    thread.completed.connect(on_completed)
    thread.error_occurred.connect(on_error)
    
    print("开始启动线程...")
    # 启动线程
    thread.start()
    print("线程已启动，等待完成...")
    
    # 运行事件循环
    app.exec_()
    print("线程已完成")
    
    # 清理临时配置文件
    if os.path.exists('thread_test_config.json'):
        os.remove('thread_test_config.json')

if __name__ == '__main__':
    test_thread()