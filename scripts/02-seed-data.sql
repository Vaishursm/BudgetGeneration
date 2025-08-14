-- Insert sample categories
INSERT OR IGNORE INTO categories (name, description) VALUES
('Technology', 'Posts about technology and programming'),
('Lifestyle', 'Posts about lifestyle and personal experiences'),
('Business', 'Posts about business and entrepreneurship'),
('Education', 'Educational content and tutorials');

-- Insert sample users (passwords are hashed versions of 'password123')
INSERT OR IGNORE INTO users (email, username, password_hash, first_name, last_name) VALUES
('admin@example.com', 'admin', '$2b$10$rOzJqQqQqQqQqQqQqQqQqO', 'Admin', 'User'),
('john@example.com', 'johndoe', '$2b$10$rOzJqQqQqQqQqQqQqQqQqO', 'John', 'Doe'),
('jane@example.com', 'janesmith', '$2b$10$rOzJqQqQqQqQqQqQqQqQqO', 'Jane', 'Smith');

-- Insert sample posts
INSERT OR IGNORE INTO posts (title, content, author_id, status) VALUES
('Welcome to Our Platform', 'This is a welcome post to introduce users to our platform. Here you can share your thoughts, ideas, and connect with others.', 1, 'published'),
('Getting Started with React', 'React is a powerful JavaScript library for building user interfaces. In this post, we will explore the basics of React development.', 2, 'published'),
('Building Full-Stack Applications', 'Learn how to build complete web applications using modern technologies like React, Node.js, and SQLite.', 2, 'draft'),
('Best Practices for Database Design', 'Database design is crucial for application performance. Here are some best practices to follow when designing your database schema.', 3, 'published');

-- Link posts to categories
INSERT OR IGNORE INTO post_categories (post_id, category_id) VALUES
(1, 2), -- Welcome post -> Lifestyle
(2, 1), -- React post -> Technology
(2, 4), -- React post -> Education
(3, 1), -- Full-stack post -> Technology
(3, 4), -- Full-stack post -> Education
(4, 1), -- Database post -> Technology
(4, 3); -- Database post -> Business
