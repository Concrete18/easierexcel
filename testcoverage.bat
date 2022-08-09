coverage run --source=easierexcel -m unittest discover
coverage html
start "" htmlcov\index.html
