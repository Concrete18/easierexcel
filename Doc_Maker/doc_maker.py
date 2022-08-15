import ast
import os


class DocMaker:
    def __init__(self, project_path):
        self.project_path = project_path

    def get_docs(self, file_path, doc):
        """
        ph
        """
        f = open(file_path, "r")
        module = ast.parse(f.read())
        class_definitions = [
            node for node in module.body if isinstance(node, ast.ClassDef)
        ]
        for class_def in class_definitions:
            # class name
            class_name = class_def.name
            if class_name:
                doc.write(f"## {class_name}\n")
            # class doc string
            doc_string = ast.get_docstring(class_def)
            if doc_string:
                doc.write(f"{doc_string}\n")
            function_definitions = [
                node for node in class_def.body if isinstance(node, ast.FunctionDef)
            ]
            for f in function_definitions:
                doc_string = ast.get_docstring(f)
                if not doc_string:
                    continue
                doc.write(f"### {f.name}\n")
                doc.write("\t" + "\t".join(doc_string.splitlines(True)))
                # print("\t---")
                # print("\t" + f.name)
                # print("\t---")
                # print("\t" + "\t".join(doc_string.splitlines(True)))

    def run(self):
        """
        ph
        """
        files = [file for file in os.listdir(self.project_path) if ".py" in file]
        print(files)
        with open("Doc_Maker\docs.md", "w") as f:
            for file in files:
                self.get_docs(os.path.join(self.project_path, file), doc=f)


App = DocMaker(project_path="easierexcel")
App.run()
