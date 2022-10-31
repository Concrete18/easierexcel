import ast, os


class DocMaker:
    """
    ph
    """

    def __init__(self, project_path, python_files=None):
        self.project_path = project_path
        self.python_files = python_files

    def get_docs(self, file_path, doc):
        """
        ph
        """
        with open(file_path, "r") as f:
            file = f.read()
        module = ast.parse(file)
        # class docs
        class_defs = [node for node in module.body if isinstance(node, ast.ClassDef)]
        for class_data in class_defs:
            # class name
            class_name = class_data.name
            if class_name:
                doc.write(f"\n\n### {class_name} Class\n")
            # class doc string
            class_doc = ast.get_docstring(class_data)
            if class_doc:
                doc.write(f"{class_doc}\n")
            # function docs
            function_defs = [
                node for node in class_data.body if isinstance(node, ast.FunctionDef)
            ]
            for f in function_defs:
                func_doc = ast.get_docstring(f)
                if not func_doc:
                    continue
                # function naming
                function_name = f.name
                doc.write(f"\n\n#### {function_name} Function\n")
                # function info
                for line in func_doc.splitlines(True):
                    # write to doc if not a TODO
                    if "TODO" not in line:
                        doc.write(line)

    def run(self):
        """
        ph
        """
        if not self.python_files:
            files = [
                file for file in os.listdir(self.project_path) if file.endswith(".py")
            ]
        else:
            files = self.python_files
        with open("Doc_Maker\docs.md", "w") as f:
            for file in files:
                file_path = os.path.join(self.project_path, file)
                self.get_docs(file_path, doc=f)


files = ["excel.py", "sheet.py"]
App = DocMaker(project_path="easierexcel", python_files=files)
App.run()
