// swift-tools-version:5.8
// The swift-tools-version declares the minimum version of Swift required to build this package.

import PythonKit

let package = Package(
    name: "Hours Tracker",
    dependencies: [
        .package(url: "https://github.com/thinkitco/PyThinkit.git", from: "1.0.0"),
    ],
    targets: [
        .executableTarget(
            name: "Hours Tracker",
            dependencies: [
                .product(name: "PyThinkit", package: "PyThinkit"),
            ],
            path: "Sources"
        ),
    ]
)
