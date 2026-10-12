import numpy as np
import matplotlib.pyplot as plt

# ============================================================
# 1. 拉普拉斯机制
# ============================================================
def laplace_mechanism(true_value, sensitivity, epsilon):
    """
    拉普拉斯机制
    :param true_value: 真实查询结果（标量或向量）
    :param sensitivity: 全局敏感度 Δf
    :param epsilon: 隐私预算 ε
    :return: 加噪后的结果
    """
    scale = sensitivity / epsilon
    noise = np.random.laplace(loc=0.0, scale=scale, size=np.shape(true_value))
    return true_value + noise


# ============================================================
# 2. 高斯机制
# ============================================================
def gaussian_mechanism(true_value, sensitivity, epsilon, delta):
    """
    高斯机制（需满足 ε-差分隐私的 (ε, δ) 版本）
    噪声标准差 σ = sqrt(2 * ln(1.25/δ)) * Δf / ε
    :param true_value: 真实查询结果
    :param sensitivity: 全局敏感度 Δf（L2 敏感度）
    :param epsilon: 隐私预算 ε
    :param delta: 失败概率 δ（通常取 1e-5 或更小）
    :return: 加噪后的结果
    """
    if delta <= 0 or delta >= 1:
        raise ValueError("delta 必须在 (0, 1) 之间")
    sigma = np.sqrt(2 * np.log(1.25 / delta)) * sensitivity / epsilon
    noise = np.random.normal(loc=0.0, scale=sigma, size=np.shape(true_value))
    return true_value + noise


# ============================================================
# 3. 示例：计数查询
# ============================================================
def demo_count_query():
    print("=" * 60)
    print("示例 1：计数查询")
    print("=" * 60)

    true_count = 1000          # 真实计数
    sensitivity = 1            # 计数查询敏感度 Δf = 1
    epsilon = 0.1              # 隐私预算
    delta = 1e-5               # 高斯机制需要

    # 拉普拉斯
    lap_result = laplace_mechanism(true_count, sensitivity, epsilon)
    print(f"[拉普拉斯] 真实值={true_count}, ε={epsilon}, Δf={sensitivity}")
    print(f"           噪声尺度 b={sensitivity/epsilon:.2f}, 加噪结果={lap_result:.2f}")

    # 高斯
    gauss_result = gaussian_mechanism(true_count, sensitivity, epsilon, delta)
    sigma = np.sqrt(2 * np.log(1.25 / delta)) * sensitivity / epsilon
    print(f"[高斯]     真实值={true_count}, ε={epsilon}, δ={delta}, Δf={sensitivity}")
    print(f"           噪声标准差 σ={sigma:.2f}, 加噪结果={gauss_result:.2f}")
    print()


# ============================================================
# 4. 示例：求和查询（数值范围 [0, 1]）
# ============================================================
def demo_sum_query():
    print("=" * 60)
    print("示例 2：求和查询（x_i ∈ [0, 1]）")
    print("=" * 60)

    data = np.random.uniform(0, 1, size=500)
    true_sum = np.sum(data)
    sensitivity = 1            # x_i ∈ [0,1] → Δf = 1
    epsilon = 0.5
    delta = 1e-5

    lap_result = laplace_mechanism(true_sum, sensitivity, epsilon)
    gauss_result = gaussian_mechanism(true_sum, sensitivity, epsilon, delta)

    print(f"真实和={true_sum:.2f}, ε={epsilon}")
    print(f"[拉普拉斯] 加噪和={lap_result:.2f}, 误差={abs(lap_result-true_sum):.2f}")
    print(f"[高斯]     加噪和={gauss_result:.2f}, 误差={abs(gauss_result-true_sum):.2f}")
    print()


# ============================================================
# 5. 示例：均值查询（拆分法：求和 + 计数）
# ============================================================
def demo_mean_query():
    print("=" * 60)
    print("示例 3：均值查询（拆分法）")
    print("=" * 60)

    data = np.random.uniform(0, 1, size=500)
    true_mean = np.mean(data)
    epsilon = 0.2
    delta = 1e-5

    # 预算对半分
    eps_sum = epsilon / 2
    eps_count = epsilon / 2

    # 求和加噪（Δf=1）
    noisy_sum = laplace_mechanism(np.sum(data), 1.0, eps_sum)
    # 计数加噪（Δf=1）
    noisy_count = laplace_mechanism(len(data), 1.0, eps_count)

    noisy_mean = noisy_sum / noisy_count
    print(f"真实均值={true_mean:.4f}, ε={epsilon}")
    print(f"加噪和={noisy_sum:.2f}, 加噪计数={noisy_count:.2f}")
    print(f"加噪均值={noisy_mean:.4f}, 误差={abs(noisy_mean-true_mean):.4f}")
    print()


# ============================================================
# 6. 可视化：噪声分布对比
# ============================================================
def plot_noise_distributions():
    sensitivity = 1.0
    epsilon = 0.5
    delta = 1e-5
    n_samples = 100000

    lap_scale = sensitivity / epsilon
    gauss_sigma = np.sqrt(2 * np.log(1.25 / delta)) * sensitivity / epsilon

    lap_noise = np.random.laplace(0, lap_scale, n_samples)
    gauss_noise = np.random.normal(0, gauss_sigma, n_samples)

    plt.figure(figsize=(10, 5))
    plt.hist(lap_noise, bins=200, density=True, alpha=0.6,
             label=f"Laplace (b={lap_scale:.2f})", range=(-20, 20))
    plt.hist(gauss_noise, bins=200, density=True, alpha=0.6,
             label=f"Gaussian (σ={gauss_sigma:.2f})", range=(-20, 20))
    plt.title(f"Noise Distribution (Δf={sensitivity}, ε={epsilon}, δ={delta})")
    plt.xlabel("Noise value")
    plt.ylabel("Density")
    plt.legend()
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig("dp_noise_distributions.png", dpi=120)
    plt.show()
    print("噪声分布图已保存为 dp_noise_distributions.png")


if __name__ == "__main__":
    np.random.seed(42)
    demo_count_query()
    demo_sum_query()
    demo_mean_query()
    plot_noise_distributions()