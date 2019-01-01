"""
该模块主要负责生成xml字符串
核心包括两个类：Element和namespace

类需要处理好名字、命名空间、属性、内容等，工厂函数主要负责方便的生成元素。
"""
from typing import List
from weakref import proxy
from collections import OrderedDict


class NameSpace:
    def __init__(self, namespace: str, alias: str=None):
        self.namespace = namespace  # type: str
        self.alias = alias  # type: str

    def __call__(self, inside):
        return NameSpaceWrapper(self, inside)

    def __repr__(self):
        if self.alias is not None:
            return f'xmlns:{self.alias}="{self.namespace}"'
        else:
            return f'xmlns="{self.namespace}"'


class NameSpaceWrapper:
    def __init__(self, namespace, inside):
        self.namespace = namespace  # type: NameSpace
        self.inside = inside  # type: str

    def __repr__(self):
        return self.attribute_name(self.inside)

    def attribute_name(self, name: str):
        return f'{self.namespace.alias}:{name}'

    def attribute_value(self):
        return self.inside


class Element:
    def __init__(self, name_: str, *children_list,
                 content: str=None, namespace_list=None,
                 **attribute_dict):
        self.name = name_  # type: str
        self.attribute_dict = OrderedDict()  # type: dict
        self.children_list = list()  # type: List[Element]
        self.content = content
        self.namespace_list = list()  # type: List[NameSpace]
        self.parent = None  # type: Element

        self.set(**attribute_dict)
        self.add(*children_list)
        if namespace_list is not None:
            self.set_ns(*namespace_list)

    def add(self, *element):
        for child in element:
            child.parent = proxy(self)
            self.children_list.append(child)
        return self

    def set(self, **attribute):
        for key, value in attribute.items():
            if isinstance(value, NameSpaceWrapper):
                key = value.attribute_name(key)
                value = value.attribute_value()
            self.attribute_dict[str(key)] = value
        return self

    def put(self, content):
        self.content = content
        return self

    def set_ns(self, *namespace):
        self.namespace_list.extend(namespace)

    def to_string(self):
        xml = f'<{self.name}'
        # processing namespace
        if len(self.namespace_list) > 0:
            for namespace in self.namespace_list:
                xml += f' {namespace}'
        # processing attribute
        for key, value in self.attribute_dict.items():
            xml += f' {key}="{value}"'
        # processing content and child element
        if len(self.children_list) > 0 or self.content is not None:
            xml += '>'
            if self.content is not None:
                xml += self.content
            for child in self.children_list:
                xml += child.to_string()
            xml += f'</{self.name}>'
        else:
            xml += '/>'

        return xml


class XML:
    def __init__(self, root: Element=None):
        self.root = root  # type: Element

    def to_string(self):
        head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        return head+self.root.to_string()


