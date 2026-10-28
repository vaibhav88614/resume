# C++ portfolio projects

These back the C++ claims on `vaibhav_cpp.tex`. **Only list a project on your resume after you have built it, run it, and can explain every line.**

## 1. mesh_threadpool (ready to build)

C++17 thread pool (mutex + condition variable + `std::packaged_task`/`std::future`), a polymorphic `MeshReader` with a Factory, and a serial-vs-parallel bounding-box benchmark.

```
cd cpp_projects/mesh_threadpool
cmake -S . -B build
cmake --build build --config Release
./build/Release/mesh_threadpool      # Windows MSVC
./build/mesh_threadpool              # Linux / MinGW
```

Interview prep questions this project lets you answer:
- Why does the worker run the task outside the lock?
- Why `notify_all` in the destructor but `notify_one` in `submit`?
- What happens if `submit` is called during shutdown?
- Why use a bounded number of chunks (`threads * 4`) instead of one task per point?
- How would you detect a data race? (try `-fsanitize=thread` on Linux)

Suggested extensions (each makes the resume stronger and truer): a real STL/VTK file reader, a VTK render of the point cloud, a unit-test target using GoogleTest, a benchmark table in this README with your real timings.

## 2. MFC viewer + COM automation (outline only - needs Windows + Visual Studio with MFC/ATL)

Not scaffolded as code because it cannot be built or verified on this machine. Suggested path:
1. Visual Studio -> MFC App (Single document) -> doc/view skeleton. Load the point cloud from project 1 into the `CDocument`, draw it in `CView::OnDraw` (or embed a VTK render window).
2. Move loading to a worker thread; post results back to the UI with `PostMessage` and a custom `WM_APP` message (never touch UI from the worker).
3. Add an ATL COM interface (`IMeshDocument` with `LoadFile`, `GetPointCount`, `GetBounds`) so Python/Excel can automate the viewer via `win32com`.
4. Write a short README with screenshots. Then list MFC/COM as hands-on instead of "familiar".
