# compilation script run through package.json

if tsc; then
    # If TypeScript compilation is successful, check if there are changes to commit
    if [ -n "$(git status --porcelain)" ]; then
        # If there are changes, add, commit, and push them
        git add .
        git commit -m 'compiled typescript'
        git push
    else
        # If there are no changes, print a message
        echo "nothing to commit"
        exit 0
    fi
else
    # If TypeScript compilation fails, print error details
    echo "compilation failed"
    exit 1
fi