# xmind_to_xls
xmind 转 excel 使用方式：

- 双击运行 MyGUI.exe

    exe 打包方式：
    
    ```shell script
    pyinstaller -Fw -i ico/Conversion.ico MyGUI.py
    ```

- 脚本运行：

    ```shell script
    python to_xls.py -xmind <path to .xmind>
    ```

结果保存到 `.xmind` 文件同级目录下的 `xls` 文件夹下。
