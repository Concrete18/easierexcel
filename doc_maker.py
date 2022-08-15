import ast
import os


class DocMaker:
    def __init__(self, project_path):
        self.project_path = project_path

    def get_docs(self, file_path):
        """
        ph
        """
        f = open(file_path, "r")
        module = ast.parse(f.read())
        class_definitions = [
            node for node in module.body if isinstance(node, ast.ClassDef)
        ]
        for class_def in class_definitions:
            print(class_def.name)
            print(ast.get_docstring(class_def))
            function_definitions = [
                node for node in class_def.body if isinstance(node, ast.FunctionDef)
            ]
            for f in function_definitions:
                doc_string = ast.get_docstring(f)
                if not doc_string:
                    continue
                print("\t---")
                print("\t" + f.name)
                print("\t---")
                print("\t" + "\t".join(doc_string.splitlines(True)))
            print("----")

    def run(self):
        """
        ph
        """
        files = [file for file in os.listdir(self.project_path) if ".py" in file]
        print(files)
        for file in files:
            self.get_docs(os.path.join(self.project_path, file))


App = DocMaker(project_path="easierexcel")
App.run()
