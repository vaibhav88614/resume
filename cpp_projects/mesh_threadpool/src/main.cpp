// Compares serial vs. thread-pool-parallel bounding-box computation on a large
// point cloud, and checks that both give the same result.

#include <chrono>
#include <cstdio>
#include <future>
#include <vector>

#include "mesh.hpp"
#include "thread_pool.hpp"

namespace {

BoundingBox computeRange(const std::vector<Point3>& pts, std::size_t begin, std::size_t end) {
    BoundingBox box;
    for (std::size_t i = begin; i < end; ++i) box.expand(pts[i]);
    return box;
}

template <class F>
double timeMs(F&& f) {
    const auto t0 = std::chrono::steady_clock::now();
    f();
    const auto t1 = std::chrono::steady_clock::now();
    return std::chrono::duration<double, std::milli>(t1 - t0).count();
}

}  // namespace

int main() {
    constexpr std::size_t kPoints = 20'000'000;

    auto reader = makeReader("synthetic", kPoints, 42);
    std::printf("Reading %zu points via '%s' reader...\n", kPoints, reader->name().c_str());
    const auto mesh = reader->read();
    const auto& pts = mesh->points;

    BoundingBox serial;
    const double serialMs = timeMs([&] { serial = computeRange(pts, 0, pts.size()); });

    ThreadPool pool;
    BoundingBox parallel;
    const double parallelMs = timeMs([&] {
        const std::size_t chunks = pool.size() * 4;
        const std::size_t step = (pts.size() + chunks - 1) / chunks;
        std::vector<std::future<BoundingBox>> futures;
        for (std::size_t b = 0; b < pts.size(); b += step) {
            const std::size_t e = std::min(b + step, pts.size());
            futures.push_back(pool.submit(computeRange, std::cref(pts), b, e));
        }
        for (auto& f : futures) parallel.merge(f.get());
    });

    const bool same = serial.min.x == parallel.min.x && serial.max.x == parallel.max.x &&
                      serial.min.y == parallel.min.y && serial.max.y == parallel.max.y &&
                      serial.min.z == parallel.min.z && serial.max.z == parallel.max.z;

    std::printf("Serial:   %.1f ms\n", serialMs);
    std::printf("Parallel: %.1f ms (%zu threads)\n", parallelMs, pool.size());
    std::printf("Results match: %s\n", same ? "yes" : "NO");
    return same ? 0 : 1;
}
