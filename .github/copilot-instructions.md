# GitHub Copilot Instructions for Python Project

## Goal

Generate high-quality, efficient, maintainable, and robust Python code following modern best practices for financial data analysis applications. Assume development is done using Visual Studio Code on Windows with Python 3.12+.

## General Guidelines

* **Language:** Python 3.12+ (leverage latest features)
* **Style:** Adhere strictly to PEP 8. Use `black` for formatting with line-length 100, and `ruff` for fast, comprehensive linting.
* **Clarity:** Prioritize readability and maintainability. Use descriptive variable/function names following domain-specific terminology (e.g., `advance_status`, `liquidation_date`, `fiscal_quarter`).
* **Type Hinting:** Use type hints for ALL function signatures, class attributes, and complex variable assignments. Leverage Python 3.12's improved type system.
* **Docstrings:** Generate Google-style docstrings for all modules, classes, functions, and methods. Include `Args:`, `Returns:`, `Raises:`, and `Examples:` sections.

## Modern Python Features (3.12+)

* **Python 3.12 Features:**
  * Type parameter syntax for generic classes and functions
  * Per-interpreter GIL for improved concurrency
  * Improved error messages with better context
  * `@override` decorator for explicit method overriding
  * Buffer protocol improvements for better performance
  * Immortal objects for better memory management
  
* **Essential Modern Features:**
  * Structural pattern matching (match/case) for complex data validation
  * The walrus operator (:=) for efficient assignment expressions
  * Dataclasses with slots=True for memory efficiency
  * Pydantic v2 models for robust data validation in financial contexts
  * F-strings with `=` for debugging (e.g., `f"{value=}"`)
  * Type unions with pipe operator (e.g., `str | None`)
  * TypeAlias, TypeGuard, ParamSpec for advanced typing
  * Self type for fluent interfaces
  * Generic TypeVar with bounds for type-safe collections

## Code Generation

* **Optimization:**
    * Leverage Python 3.12's performance improvements (faster startup, improved memory usage)
    * Use built-in functions and standard library modules optimized for 3.12
    * Prefer list/dict/set comprehensions with conditional expressions
    * Use generator expressions for memory-efficient iterations
    * Implement `__slots__` in classes for better memory usage
    * Use `functools.cache` instead of `lru_cache` for unlimited caching
    * Consider using `memoryview` for large data operations
    * Be mindful of algorithmic complexity - document Big O notation for complex functions
    
* **Python 3.12 Specific Optimizations:**
    * Use the new type parameter syntax for cleaner generic code
    * Leverage improved error messages for better debugging
    * Use `@override` decorator to prevent inheritance bugs
    * Take advantage of per-interpreter GIL for true parallelism
    * Use immortal objects for frequently used constants
    
* **Accuracy:**
    * Pay close attention to the surrounding code context and comments
    * Generate code that leverages existing patterns in the codebase
    * If requirements are ambiguous, provide a clear implementation with documented alternatives
    * Include type hints that work with Python 3.12's enhanced type system
    * Use domain-specific terminology consistently

## Directory Structure

Assume and maintain the following project structure:

```
project-root/
├── .github/
│   └── copilot-instructions.md
├── src/
│   └── advance_analysis/
│       ├── __init__.py
│       ├── main.py
│       ├── core/
│       ├── modules/
│       └── utils/
├── tests/
│   ├── __init__.py
│   └── test_*.py
├── docs/
│   └── ...
├── data/          # Optional: For data files
├── scripts/       # Optional: For helper scripts
├── .gitignore
├── pyproject.toml # Or requirements.txt
└── README.md
```

* Place core application logic within `src/advance_analysis/`.
* Place unit and integration tests within `tests/`. Test files should mirror the structure of the `src/` directory.
* Use relative imports within the `src/` directory (e.g., `from .core import ...`).

## Dependency Management

* **Dependency Management:**
  * Use Poetry for comprehensive dependency management
  * Separate dev dependencies from production dependencies
  * Pin dependencies with specific versions for reproducibility
  * Consider using a tool like `pip-compile` for generating deterministic requirements.txt
  * Use `pipdeptree` to visualize and manage complex dependency trees
  * Implement dependency security scanning in CI pipeline

## Virtual Environment Management

* **Environment Management:**
  * Use venv or virtualenv for isolated environments
  * Consider using pyenv for Python version management
  * Document environment setup steps in README
  * Create environment setup scripts if complex
  * Use .env files with python-dotenv for environment variables
  * Consider creating a dev container configuration for VSCode

## Error Handling

* **Error Handling Strategy:**
  * Catch specific exceptions rather than generic `Exception`.
  * Define custom exception classes for application-specific errors when appropriate.
  * Use `try...except...finally` blocks for cleanup operations (e.g., closing files or network connections).
  * Use context managers (`with` statement) for resource management (files, locks, connections).
  * Provide clear and informative error messages.
  * Design an error hierarchy specific to application domains
  * Implement error codes for systematic troubleshooting
  * Create appropriate recovery mechanisms for different error types
  * Log contextual information with exceptions
  * Consider using exception chaining with `raise ... from ...` syntax

## Debugging

* **Variable Names:** Use descriptive variable names.
* **Intermediate Variables:** Don't shy away from using intermediate variables to clarify steps in complex calculations or logic.
* **Pure Functions:** Prefer pure functions (functions whose output depends only on their input and have no side effects) where possible, as they are easier to test and debug.
* **Assertions:** Use `assert` statements for sanity checks during development (but be aware they can be disabled).
* **Debugging Tools:**
  * Utilize VSCode's built-in debugger with breakpoints
  * Consider using pdb/ipdb for command-line debugging
  * Use logging for persistent debugging information
  * Add debugging decorators for function entry/exit tracing
  * Implement custom debug views for complex data structures

## Logging

* **Standard Library:** Use the built-in `logging` module.
* **Configuration:** Configure logging early in the application's entry point (`main.py`). Consider configuration via file (`logging.conf`) or dictionary (`logging.config.dictConfig`).
* **Levels:** Use appropriate logging levels (`DEBUG`, `INFO`, `WARNING`, `ERROR`, `CRITICAL`).
* **Context:** Include relevant context in log messages (e.g., function names, relevant variable values).
* **Avoid `print()`:** Replace `print()` statements used for debugging or status updates with appropriate `logging` calls.
* **Advanced Logging:**
  * Implement structured logging with JSON formatter for machine parsing
  * Use log rotation to manage log file sizes
  * Consider centralized logging for distributed applications
  * Add correlation IDs for tracking requests across components
  * Use custom log handlers for specific log destinations

## Testing Strategy

* **Testing Framework:**
  * Use `pytest` 7.0+ with Python 3.12 compatibility
  * Leverage pytest's improved assertion introspection in Python 3.12
  * Generate tests that aim for high code coverage (90%+ for critical financial logic)
  * Use `pytest` fixtures with proper scope management (function, class, module, session)
  * Implement fixture factories for complex test data generation
  * Use `pytest-asyncio` for testing async code
  * Use `unittest.mock` with `spec=True` for type-safe mocking
  * Test functions should have descriptive names: `test_<function>_<scenario>_<expected_result>`

* **Financial Application Testing:**
  * **Data Validation Tests:**
    * Test boundary conditions for financial calculations
    * Validate decimal precision handling
    * Test date range validations (fiscal years, quarters)
    * Verify Excel file format compatibility
  * **Business Logic Tests:**
    * Test advance status transitions
    * Validate liquidation date calculations
    * Test compliance rule engines
    * Verify report generation accuracy
  * **GUI Testing:**
    * Use pytest-qt or similar for GUI component testing
    * Test threading behavior and UI responsiveness
    * Validate user input handling and error displays
    * Test file dialog operations

* **Comprehensive Testing:**
  * Unit tests with mocked dependencies
  * Integration tests with test databases/files
  * End-to-end tests simulating user workflows
  * Property-based testing with hypothesis for data transformations
  * Snapshot testing for report outputs
  * Performance benchmarks for large dataset processing
  * Parameterized tests using `pytest.mark.parametrize`
  * Fuzzing tests for input validation

* **Test Quality:**
  * Follow AAA pattern with clear sections
  * Use `pytest.raises` for exception testing
  * Implement custom assertions for domain logic
  * Use test coverage with branch coverage enabled
  * Create test data factories for consistent test data
  * Document complex test scenarios
  * Use `pytest.mark` for test categorization (slow, integration, unit)
  * Implement continuous test monitoring

## Documentation

* **Documentation Tools:**
  * Use Sphinx or MkDocs for generating documentation
  * Include code examples in docstrings to demonstrate usage
  * Consider using autodoc extensions to generate API docs from docstrings
  * Create a detailed README.md with installation, usage, and contribution guidelines
  * Generate online documentation on each release
* **Documentation Content:**
  * Include architecture diagrams (consider using Mermaid or PlantUML)
  * Document design decisions and trade-offs
  * Create user guides with examples
  * Maintain API documentation with examples
  * Document configuration options and environment variables

## Security

* **Security Practices:**
  * Always validate and sanitize external input (user input, API responses, file contents).
  * Do not hardcode secrets (API keys, passwords). Use environment variables or a dedicated secrets management tool.
  * Suggest placeholders like `os.getenv("API_KEY")` or use python-dotenv.
  * Be mindful of third-party library security vulnerabilities.
  * Implement input validation with pydantic or marshmallow
  * Use parameterized queries for database access
  * Conduct regular dependency scanning with safety
  * Follow OWASP guidelines for web applications
  * Implement content security policies where appropriate
  * Use secure hashing and password storage (argon2, bcrypt)

## Code Quality Automation

* **Pre-commit Hooks:**
  * Set up pre-commit hooks for:
    * Code formatting (black)
    * Import sorting (isort)
    * Linting (ruff/flake8)
    * Type checking (mypy)
    * Security scanning (bandit)
  * Include .pre-commit-config.yaml in project
* **CI/CD Integration:**
  * Set up GitHub Actions workflows for automated testing
  * Configure linting and type checking in CI pipelines
  * Implement automated deployment processes
  * Set up code coverage reporting
  * Implement parallel testing for faster feedback
  * Create deployment pipelines for different environments

## Static Type Checking

* **Type Checking Tools:**
  * Use mypy for static type checking
  * Configure mypy.ini or pyproject.toml with appropriate strictness
  * Consider runtime type validation with libraries like Pydantic
  * Use Protocol classes for duck typing
  * Implement TypeGuard for runtime type narrowing
  * Use TypeVar for generic function definitions
  * Create custom types for domain-specific concepts

## Performance Optimization

* **Performance Analysis:**
  * Profile code with cProfile or py-spy
  * Use line_profiler for line-by-line profiling
  * Consider memory_profiler for memory usage analysis
  * Implement benchmarking tests for critical paths
  * Monitor performance metrics in production
* **Optimization Techniques:**
  * Use caching strategically (functools.lru_cache, redis)
  * Implement batch processing for I/O operations
  * Consider using compiled extensions for performance-critical sections
  * Parallelize CPU-bound operations with multiprocessing
  * Optimize database queries with proper indexing

## Asynchronous Programming

* **Async Development:**
  * Use asyncio for I/O-bound operations
  * Prefer async/await syntax over callbacks
  * Consider libraries like aiohttp for async HTTP
  * Use asyncio.gather for concurrent tasks
  * Implement proper error handling in async code
  * Be mindful of event loop blocking
  * Consider using async frameworks for web applications

## Containerization

* **Container Strategy:**
  * Provide Dockerfile for consistent environment
  * Use multi-stage builds for optimized containers
  * Create docker-compose.yml for local development
  * Document container usage in README
  * Optimize container image size
  * Include health checks in container configurations
  * Implement proper signal handling for graceful shutdowns

## Cross-Platform Compatibility

* **Platform Independence:**
  * Use pathlib for file path manipulation
  * Avoid platform-specific commands or libraries
  * Test on multiple platforms when possible
  * Use CI with multiple OS runners
  * Handle line ending differences properly
  * Be mindful of file permission differences between platforms
  * Use platform-independent temporary file handling

## Modern GUI Development

* **Framework Selection:**
  * Primary: CustomTkinter for modern, themed tkinter applications
  * Alternative: Kivy or PyQt6 for more complex requirements
  * Web-based: Streamlit or Gradio for data analysis dashboards
  * Consider Flet for Flutter-based modern UIs

* **GUI Best Practices:**
  * **Responsive Design:**
    * Use grid and pack managers effectively
    * Implement window resizing with proper widget scaling
    * Support multiple screen resolutions and DPI settings
    * Use relative sizing (percentages) over fixed pixels
  
  * **User Experience (UX):**
    * Implement progress bars for long-running operations
    * Use threading/asyncio to prevent UI freezing
    * Provide clear, actionable error messages in dialogs
    * Include tooltips for complex controls
    * Implement keyboard shortcuts for power users
    * Add context menus for quick actions
    * Use consistent color schemes and fonts
    * Implement dark/light theme switching
  
  * **Accessibility:**
    * Support keyboard navigation (Tab order)
    * Use high contrast colors for readability
    * Include screen reader support where possible
    * Provide text alternatives for icons
    * Ensure minimum clickable area sizes (44x44 pixels)
  
  * **Modern Features:**
    * Drag-and-drop file selection
    * Auto-complete for text inputs
    * Real-time validation with visual feedback
    * Collapsible sections for complex forms
    * Search/filter capabilities for data tables
    * Export functionality (PDF, Excel, CSV)
    * Undo/Redo functionality for critical actions
    * Auto-save and recovery features

* **GUI Architecture:**
  * **MVC/MVP Pattern:**
    * Separate business logic from UI code
    * Use observer pattern for data binding
    * Implement command pattern for actions
    * Create reusable custom widgets
  
  * **State Management:**
    * Centralize application state
    * Use event-driven architecture
    * Implement proper state persistence
    * Handle concurrent state updates safely
  
  * **Performance:**
    * Use virtual scrolling for large lists
    * Implement lazy loading for data
    * Cache frequently accessed resources
    * Debounce user input events
    * Use background threads for I/O operations

## Project-Specific Guidance

* **Financial Data Processing:**
  * Use pandas with proper dtype specifications for memory efficiency
  * Implement decimal precision for monetary calculations
  * Use numpy for statistical operations
  * Consider polars for faster data processing
  * Validate data integrity at every transformation step
  * Implement audit trails for data modifications
  * Use openpyxl for Excel file operations with formatting preservation
  * Handle date/time with timezone awareness
  * Implement fiscal year/quarter calculations correctly

* **Excel Integration:**
  * Preserve Excel formatting when reading/writing
  * Handle merged cells and formulas appropriately
  * Implement sheet-level operations efficiently
  * Support Excel table structures
  * Maintain cell styles and conditional formatting
  * Handle large Excel files with streaming APIs

* **Validation & Compliance:**
  * Implement multi-level validation strategies
  * Create detailed validation reports
  * Support configurable business rules
  * Maintain validation history
  * Generate compliance documentation
  * Implement role-based access controls

## Data Quality & Financial Best Practices

* **Data Integrity:**
  * Implement checksums for critical data transfers
  * Use transaction patterns for multi-step operations
  * Maintain audit logs with timestamps and user info
  * Implement data versioning for tracking changes
  * Use immutable data structures where appropriate
  * Validate data at ingestion, transformation, and output stages

* **Financial Calculations:**
  * Use `decimal.Decimal` for monetary values, never float
  * Implement proper rounding rules (ROUND_HALF_EVEN for banking)
  * Handle currency with explicit currency codes
  * Implement fiscal year/quarter logic correctly
  * Use business day calculations for date operations
  * Handle timezone conversions for global operations

* **Excel Best Practices:**
  * Preserve formulas and formatting when possible
  * Handle Excel's date system quirks (1900 vs 1904)
  * Implement proper error handling for corrupted files
  * Use streaming for large files to manage memory
  * Validate sheet names and handle special characters
  * Implement cell validation rules

* **Compliance & Audit:**
  * Log all data modifications with user context
  * Implement data retention policies
  * Create reproducible analysis pipelines
  * Generate detailed validation reports
  * Implement access controls and data masking
  * Maintain change history for critical configurations

## Version Control Best Practices

* **Git Workflow:**
  * Use descriptive branch names (feature/, bugfix/, hotfix/)
  * Follow conventional commits: `type(scope): description`
  * Keep commits focused and atomic
  * Use Pull Requests with required reviews
  * Implement branch protection rules
  * Tag releases with semantic versioning
  * Use git hooks for pre-commit validation
  * Maintain a clean commit history with rebase