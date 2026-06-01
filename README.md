# HSC Sort (Holographic Square Constructor Sort)

An official repository containing the implementation, test environment, and interactive examples of the **SC Sort (Square Constructor Sort)** and **HSC Sort (Holographic Square Constructor Sort)** algorithms.

## 📌 Project Overview
HSC Sort is a hybrid parallel sorting algorithm based on the geometric properties of the Steiner system of pairs $S(2, k, k^2)$ and combination sets denoted as $C[\text{pair}, k(\text{rows}), n(k^2)]$. By combining row sorting with column rotation (*slicing*) into a single composite cycle, the algorithm provides an ideal foundation for deployment on $n$ parallel processors.

### Key Advantages:
* **Deterministic Execution Time:** In Holographic mode (HG), the internal complexity of sub-sorting algorithms is completely absorbed, linearizing the overall execution time.
* **Redundancy & Sparse Data Immunity:** Due to integrated row-wise sorting, the architecture is highly resilient against massive data duplication and incomplete/sparse matrices.
* **Low Latency:** Parallel processes run without synchronization conflicts during the core execution cycles.

---

## 📂 Repository Structure

The project is organized into three main directories:

* `Demo/` — Contains interactive spreadsheets and visual examples of the constructor functions ($C(2,3,9)$, $C(2,4,16)$, etc.), mapping the row sorting and column rotation cycles.
* `Generator/` — Contains the main benchmark workbench `Comparative HSC Sort.ods` (with integrated macros) along with `Module.py`, which holds the identical Python source code loaded within the spreadsheet for standalone reference.
* `Test Protocols/` — Contains the evaluation logs and benchmarking benchmarks exported into LibreOffice Writer text documents mapping execution times, communication overheads, and parallel efficiency.
* `README.md` — This documentation file.

---

## 📊 Interactive Spreadsheet Examples (`Demo of constructor functions.xlsx`)
To fully grasp the geometric and algebraic nature of the algorithm, this repository includes an interactive spreadsheet with 3 to 5 clear visual examples mapping the transformation process. 

The examples showcase the exact progression of **Row Sorting + Column Rotation** over $(n+1)$ cycles. 

### 💡 Note on Combinatorial Configurations & Latin Squares:
In pure mathematics, constructing specific block designs often collides with the well-known challenges of **Latin Squares** (their existence, orthogonality, and constraints for certain even/odd dimensions). 

**Important for users:** While this represents a significant bottleneck for strict algebraic constructions, **it does not affect the SC/HSC Sort algorithm**. The underlying architecture successfully processes both even and odd dimensions, as well as incomplete matrices. Any structural asymmetry or padding requirement is fully resolved by the integrated row-sorting phase and final merge operations, making the sorting mechanism universally applicable.

---

## ⚙️ How to Run the Benchmarks
The benchmarking scripts are implemented in pure Python and intentionally **do not rely on low-level optimized libraries like NumPy**. This ensures an objective evaluation of the pure algorithmic and structural logic of the matrix constructor.

### Prerequisites:
* Python 3.x
* A spreadsheet processor (LibreOffice Calc or Microsoft Excel) to view the interactive sheets and protocols.

### Execution:
Simply run the python script via terminal:
```bash
python Module.py
```

## 📜 License
This project is licensed under the MIT License - see the LICENSE file for details.

