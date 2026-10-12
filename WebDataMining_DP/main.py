import numpy as np
from pure_ldp.frequency_oracles.unary_encoding import UEClient, UEServer

# 模拟数据：10000 个用户，每个用户有一个查询词 ID（1-1000）
epsilon = 3.0
d = 1000  # 查询词总数

# 创建客户端和服务器
client = UEClient(epsilon=epsilon, d=d, use_oue=True)
server = UEServer(epsilon=epsilon, d=d, use_oue=True)

# 真实数据：Zipf 分布
np.random.seed(42)
ranks = np.arange(1, d + 1)
probs = 1 / ranks
probs = probs / probs.sum()
data = np.random.choice(d + 1, size=10000, p=probs)  # 值为 1 到 d

# 客户端扰动 + 服务器聚合
for item in data:
    priv = client.privatise(item)
    server.aggregate(priv)

# 估计频率
estimates = server.estimate_all(range(1, d + 1))
print("前 10 个查询词的真实频率和估计频率：")
for i in range(10):
    print(f"查询 {i+1}: 真实 ≈ {np.sum(data == i+1)}, 估计 ≈ {estimates[i]}")