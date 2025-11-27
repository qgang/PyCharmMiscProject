import pandas as pd

if __name__ == "__main__":
    data = {
        '学生姓名': ['张三', '李四', '王五'],
        '数学': [85, 90, 78.23],
        '英语': [88, 92, 85],
        '班级': [1, 1, 2]
    }
    df = pd.DataFrame(data)
    class_avg = df.groupby('班级')['数学'].mean().round(1)
    print(class_avg)

