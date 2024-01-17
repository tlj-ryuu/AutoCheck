now=$(date "+%Y-%m-%d %H:%M:%S")
echo "Change Directory to D:/project_autocheck"
cd D:/project_autocheck
echo "Starting add-commit-pull-push..."
git add . && git commit -m "$now" && git pull origin main && git push origin master:main
echo "Done!"