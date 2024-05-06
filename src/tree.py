from typing import List, Optional, TypedDict, Union, Generator

from office365.sharepoint.files.file import File
from office365.sharepoint.folders.folder import Folder


ChildNode = Union["FileNode", "FolderNode"]
ChildNodeDict = Union["FileNodeDict", "FolderNodeDict"]


class FolderNodeDict(TypedDict):
    name: str
    path: str
    type: str
    children: List[ChildNodeDict]


class FileNodeDict(TypedDict):
    name: str
    path: str
    type: str


class Node:
    def __init__(
        self,
        obj: Union[File, Folder],
        type: str,
        parent: Optional["FolderNode"],
    ) -> None:
        self.obj = obj
        self.name = obj.properties["Name"]
        self.path = obj.properties["ServerRelativeUrl"]
        self.type = type
        self.parent = parent

    def set_parent(self, parent: "FolderNode") -> None:
        self.parent = parent

    def is_file(self) -> bool:
        return self.type == "file"

    def is_folder(self) -> bool:
        return self.type == "folder"

    def __str__(self) -> str:
        class_name = self.__class__.__name__
        return f"{class_name}(name='{self.name}')"

    def __repr__(self) -> str:
        return self.__str__()


class FolderNode(Node):
    def __init__(
        self,
        obj: Folder,
        parent: Optional["FolderNode"] = None,
        children: Optional[List[ChildNode]] = None,
    ) -> None:
        super().__init__(obj=obj, type="folder", parent=parent)
        self.children = children or []

    def add_child(self, child: ChildNode) -> None:
        child.set_parent(self)
        self.children.append(child)

    def to_dict(
        self,
    ) -> FolderNodeDict:
        return FolderNodeDict(
            name=self.name,
            path=self.path,
            type=self.type,
            children=[child.to_dict() for child in self.children],
        )


class FileNode(Node):
    def __init__(self, obj: File, parent: FolderNode) -> None:
        super().__init__(obj=obj, type="file", parent=parent)

    def to_dict(self) -> FileNodeDict:
        return FileNodeDict(name=self.name, path=self.path, type=self.type)


class Tree:
    def __init__(self, root: FolderNode) -> None:
        self.root = root
        self.depth = 0

    def set_current_node(self, folder_node: FolderNode) -> None:
        self.current_folder_node = folder_node
        self.depth += 1

    def to_dict(self) -> FolderNodeDict:
        return self.root.to_dict()

    def __iter__(self) -> Generator[ChildNode, None, None]:
        stack: List[ChildNode] = [self.root]
        while stack:
            node = stack.pop()
            yield node
            if isinstance(node, FolderNode):
                stack.extend(reversed(node.children))

    def __str__(self) -> str:
        return f"Tree(root={self.root} depth={self.depth})"

    def __repr__(self) -> str:
        return self.__str__()
