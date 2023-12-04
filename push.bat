@echo off

REM 检查是否有修改
git status --porcelain "1.txt" >nul && (
    REM 添加所有修改
    git add .

    REM 提交修改
    git commit -m "Auto commit: %date% %time%"

    REM 推送到远程仓库（假设远程仓库名为 origin，主分支名为 main）
    git push self main
) || (
    echo No changes to commit.
)

pause