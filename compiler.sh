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
    fi
else
    # If TypeScript compilation fails, print error details
    echo "compilation failed"
fi