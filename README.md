# Excel Tool

### **Requires**
- Python 3+
- xlrd 1.2.0

### **Install xlrd**
```c#
pip install xlrd==1.2.0
```

### **Features**
- Convert excel to
  - json
  - lua
  - js/ts
  - ~~protobuf(**TODO**)~~
- Convert scripts
  - c#
  - ~~other(TODO)~~
- Multi-process



### **Excel structure**

**Cell(0,0)**
- object
- array
- dict
- group


| object|       |              |        |
| ----- | ----- | ------------ | ------ |
| desc1 | attr1 | int          | 1      |
| desc2 | attr2 | bool         | 0      |
| desc3 | attr3 | int16        | 3      |
| desc3 | attr4 | int32        | 3      |
| desc3 | attr5 | float        | 3.0    |
| desc3 | attr6 | double       | 3.0    |
| desc3 | attr7 | string       | abc    |
| desc4 | attr9 | int[]        |[1,2,3] |
| desc4 | attr0 | int[][]      |[[1,2,3],[1,2,3]] |

```JSON
{
    "attr1": 1,
    "attr2": 0,
    "attr3": 3,
    "attr4": 3,
    "attr5": 3.0,
    "attr6": 3.0,
    "attr7": "abc",
    "attr8": [1,2,3],
    "attr9": [1,2,3],
    "attr0": [[1,2,3],[1,2,3]]
}
```


| array | desc | desc  | desc   |
| ----- | ---- | ----- | ------ |
| id    | att1 | att2  | att3   |
| int   | bool | float | double |
| 1     | false| 1.0   | 1.0    |
| 2     | false| 2.0   | 2.0    |
| 3     | false| 3.0   | 3.0    |

```JSON
[
    {"id":1, "attr1":false, "attr2":1.0, "attr3":1.0},
    {"id":2, "attr1":false, "attr2":2.0, "attr3":2.0},
    {"id":3, "attr1":false, "attr2":3.0, "attr3":3.0}
]
```


| dict  | desc | desc  | desc   |
| ----- | ---- | ----- | ------ |
| id    | att1 | att2  | att3   |
| int   | bool | float | double |
| 1     | false| 1.0   | 1.0    |
| 2     | false| 2.0   | 2.0    |
| 3     | false| 3.0   | 3.0    |

```JSON
{
    "1": {"id":1, "attr1":false, "attr2":1.0, "attr3":1.0},
    "2": {"id":2, "attr1":false, "attr2":2.0, "attr3":2.0},
    "3": {"id":3, "attr1":false, "attr2":3.0, "attr3":3.0}
}
```

| group | desc | desc  | desc   |
| ----- | ---- | ----- | ------ |
| id    | att1 | att2  | att3   |
| int   | bool | float | double |
| 1     | false| 1.0   | 1.0    |
| 1     | false| 1.0   | 1.0    |
| 1     | false| 1.0   | 1.0    |
| 2     | false| 2.0   | 2.0    |
| 2     | false| 2.0   | 2.0    |
| 3     | false| 3.0   | 3.0    |

```JSON
{
    "1":
    [
        {"id":1, "attr1":false, "attr2":1.0, "attr3":1.0},
        {"id":1, "attr1":false, "attr2":1.0, "attr3":1.0},
        {"id":1, "attr1":false, "attr2":1.0, "attr3":1.0}
    ],
    "2":
    [
        {"id":2, "attr1":false, "attr2":2.0, "attr3":2.0},
        {"id":2, "attr1":false, "attr2":2.0, "attr3":2.0}
    ],
    "3":
    [
        {"id":3, "attr1":false, "attr2":3.0, "attr3":3.0}
    ]
}
```