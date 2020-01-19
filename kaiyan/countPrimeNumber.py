# 定义一个数组，用来存放质数
num = []
# 计算出2到100的质数
for i in range(2, 100):
    # j是除数，用i去除于比i小的所有数
    for j in range(2, i):
        # 如果i被j整除，跳出j这个循环（表示这个i不是质数，可以算下一个)
        if i % j == 0:
            break
    # j循环处理完，没有跳出循环，j从2走到i-1的数就是质数
    else:
        # 把这个数放到num这个数组
        num.append(i)
# 质数计算完毕后，输出数组
print(num)


# 知识点

# for i in range(a, b):
# i从a走到b-1

# for else
# 《你上网查查，答案填在这里》

# break else
# 《你上网查查，答案填在这里》
